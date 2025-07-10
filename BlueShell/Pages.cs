using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.ComponentModel;
using System.Linq.Expressions;

public class ComTab : TabPage
{
    // Windows API 导入
    [DllImport("user32.dll")]
    private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
    private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
    private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
    private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
    private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetFocus();

    // 常量定义
    private const int GWL_STYLE = -16;
    private const int WS_CAPTION = 0x00C00000;
    private const int WS_THICKFRAME = 0x00040000;
    private const int MAX_WAIT_TIME = 10000; // 10秒超时
    private const int CHECK_INTERVAL = 500; // 检查间隔(毫秒)

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    private Process _externalProcess;
    private IntPtr _externalWindowHandle = IntPtr.Zero;
    private string _processName;
    private System.Threading.Timer _monitorTimer;
    private TabControl _parentTabControl;
    private bool _isDisposing = false;

    // 处理32位和64位系统的WindowLong函数
    private static int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong)
    {
        if (IntPtr.Size == 4)
            return SetWindowLong32(hWnd, nIndex, dwNewLong);
        else
            return (int)SetWindowLongPtr64(hWnd, nIndex, new IntPtr(dwNewLong));
    }

    private static int GetWindowLong(IntPtr hWnd, int nIndex)
    {
        if (IntPtr.Size == 4)
            return GetWindowLong32(hWnd, nIndex);
        else
            return (int)GetWindowLongPtr64(hWnd, nIndex);
    }

    public ComTab(string tabName, string programCommand) : base(tabName)
    {
        try
        {
            _processName = programCommand;

            // 启动外部程序
            StartExternalProcess(programCommand);

            // 查找窗口句柄
            FindExternalWindowHandle();

            if (_externalWindowHandle == IntPtr.Zero)
            {
                throw new Exception("无法找到外部程序窗口句柄");
            }

            // 将外部窗口嵌入到当前TabPage中
            EmbedExternalWindow();

            // 启动监控定时器
            StartMonitoring();

            // 设置焦点到嵌入窗口
            SetFocusToEmbeddedWindow();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"创建ComTab失败: {ex.Message}");
            DisposeResources();
            throw;
        }
    }

    protected override void OnParentChanged(EventArgs e)
    {
        base.OnParentChanged(e);

        // 获取父TabControl引用
        _parentTabControl = this.Parent as TabControl;
    }

    private void StartMonitoring()
    {
        // 创建定时器检查窗口状态
        _monitorTimer = new System.Threading.Timer(CheckWindowStatus, null, CHECK_INTERVAL, CHECK_INTERVAL);
    }

    private void CheckWindowStatus(object state)
    {
        if (_isDisposing) return;

        if (this.InvokeRequired)
        {
            this.Invoke(new Action<object>(CheckWindowStatus), state);
            return;
        }

        try
        {
            bool shouldRemove = false;

            // 检查进程状态
            if (_externalProcess != null && _externalProcess.HasExited)
            {
                shouldRemove = true;
            }
            // 检查窗口句柄是否有效
            else if (_externalWindowHandle != IntPtr.Zero && !IsWindow(_externalWindowHandle))
            {
                shouldRemove = true;
            }

            if (shouldRemove)
            {
                RemoveSelfFromParent();
            }
            else
            {
                // 定期恢复光标焦点
                RestoreCursorFocus();
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"监控窗口状态时出错: {ex.Message}");
        }
    }

    private void SetFocusToEmbeddedWindow()
    {
        if (_externalWindowHandle != IntPtr.Zero)
        {
            SetForegroundWindow(_externalWindowHandle);
            SetFocus(_externalWindowHandle);
        }
    }

    private void RestoreCursorFocus()
    {
        if (_externalWindowHandle != IntPtr.Zero && IsWindow(_externalWindowHandle))
        {
            // 检查当前焦点是否在嵌入窗口中
            IntPtr focusedWindow = GetFocus();
            if (focusedWindow != _externalWindowHandle)
            {
                SetFocus(_externalWindowHandle);
            }
        }
    }

    private void RemoveSelfFromParent()
    {
        if (_isDisposing) return;
        _isDisposing = true;

        if (_parentTabControl != null && _parentTabControl.TabPages.Contains(this))
        {
            _parentTabControl.TabPages.Remove(this);
        }

        // 清理资源
        DisposeResources();
    }

    private void StartExternalProcess(string programCommand)
    {
        _externalProcess = new Process();
        _externalProcess.EnableRaisingEvents = true;
        _externalProcess.Exited += ExternalProcess_Exited;

        // 处理带参数的命令
        if (programCommand.Contains(" "))
        {
            int firstSpace = programCommand.IndexOf(' ');
            _externalProcess.StartInfo.FileName = programCommand.Substring(0, firstSpace);
            _externalProcess.StartInfo.Arguments = programCommand.Substring(firstSpace + 1);
        }
        else
        {
            _externalProcess.StartInfo.FileName = programCommand;
        }

        _externalProcess.StartInfo.UseShellExecute = true;
        _externalProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal;
        try
        {


            _externalProcess.Start();
        }
        catch (Exception ex) 
        {
            MessageBox.Show(ex.ToString(), "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            Environment.Exit(1);
        }
    }

    private void ExternalProcess_Exited(object sender, EventArgs e)
    {
        if (_isDisposing) return;

        if (this.InvokeRequired)
        {
            this.Invoke(new Action<object, EventArgs>(ExternalProcess_Exited), sender, e);
            return;
        }

        RemoveSelfFromParent();
    }

    private void FindExternalWindowHandle()
    {
        DateTime startTime = DateTime.Now;
        bool windowFound = false;

        // 等待窗口出现，最多等待MAX_WAIT_TIME毫秒
        while (!windowFound && (DateTime.Now - startTime).TotalMilliseconds < MAX_WAIT_TIME)
        {
            // 首先尝试获取主窗口句柄
            _externalProcess.Refresh();
            _externalWindowHandle = _externalProcess.MainWindowHandle;

            if (_externalWindowHandle != IntPtr.Zero)
            {
                windowFound = true;
                break;
            }

            // 如果主窗口句柄不可用，枚举所有窗口查找属于该进程的可见窗口
            uint processId = (uint)_externalProcess.Id;
            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                uint windowProcessId;
                GetWindowThreadProcessId(hWnd, out windowProcessId);

                if (windowProcessId == processId && IsWindowVisible(hWnd))
                {
                    _externalWindowHandle = hWnd;
                    windowFound = true;
                    return false; // 停止枚举
                }
                return true; // 继续枚举
            }, IntPtr.Zero);

            if (!windowFound)
            {
                Thread.Sleep(100); // 等待100ms再试
            }
        }

        if (!windowFound)
        {
            throw new Exception($"在{MAX_WAIT_TIME / 1000}秒内未找到程序窗口");
        }
    }

    private void EmbedExternalWindow()
    {
        if (_externalWindowHandle == IntPtr.Zero)
            return;

        // 移除窗口的标题栏和边框
        int style = GetWindowLong(_externalWindowHandle, GWL_STYLE);
        style = style & ~WS_CAPTION & ~WS_THICKFRAME;
        SetWindowLong(_externalWindowHandle, GWL_STYLE, style);

        // 设置父窗口为当前TabPage
        SetParent(_externalWindowHandle, this.Handle);

        // 调整窗口大小以适应TabPage
        ResizeEmbeddedWindow();
    }

    private void ResizeEmbeddedWindow()
    {
        if (_externalWindowHandle == IntPtr.Zero)
            return;

        // 获取TabPage的客户区大小
        int width = this.ClientSize.Width;
        int height = this.ClientSize.Height;

        // 调整嵌入窗口的大小和位置
        MoveWindow(_externalWindowHandle, 0, 0, width, height, true);
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        ResizeEmbeddedWindow();
    }

    private void DisposeResources()
    {
        if (_isDisposing) return;
        _isDisposing = true;

        // 停止监控定时器
        if (_monitorTimer != null)
        {
            _monitorTimer.Dispose();
            _monitorTimer = null;
        }

        if (_externalProcess != null)
        {
            try
            {
                _externalProcess.Exited -= ExternalProcess_Exited;

                if (!_externalProcess.HasExited)
                {
                    // 先尝试友好关闭
                    if (_externalWindowHandle != IntPtr.Zero)
                    {
                        // 恢复窗口父关系
                        SetParent(_externalWindowHandle, IntPtr.Zero);
                        // 恢复窗口样式
                        int style = GetWindowLong(_externalWindowHandle, GWL_STYLE);
                        style = style | WS_CAPTION | WS_THICKFRAME;
                        SetWindowLong(_externalWindowHandle, GWL_STYLE, style);
                        // 调整窗口大小
                        MoveWindow(_externalWindowHandle, 0, 0, 800, 600, true);
                    }

                    _externalProcess.CloseMainWindow();

                    if (!_externalProcess.WaitForExit(2000)) // 等待2秒
                    {
                        _externalProcess.Kill();
                    }
                }
                _externalProcess.Dispose();
            }
            catch
            {
                // 忽略清理过程中的异常
            }
            finally
            {
                _externalProcess = null;
                _externalWindowHandle = IntPtr.Zero;
            }
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            DisposeResources();
        }
        base.Dispose(disposing);
    }
}