using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.String;

namespace BlueShell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void cmd_enterd() 
        {
            string cmd = textBox1.Text.Split(' ')[0];
            string[] arg = (textBox1.Text.Length != cmd.Length) ? textBox1.Text.Substring(cmd.Length + 1).Split(' ') : null;
            ;
            switch (cmd)
            {
                case "new":
                    NewWindow(arg);
                    break;
            }
        }
        private void NewWindow(string[] arg) 
        {
            TabPage next;
            if (arg == null) 
            {
                MessageBox.Show("please enter window type");
                
            }

            next = new ComTab(arg[0],Join(" ",arg));
            MainTC.TabPages.Add(next);
        }

        
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;
            if (e.KeyCode == Keys.Enter) cmd_enterd();
        }
    }
}
