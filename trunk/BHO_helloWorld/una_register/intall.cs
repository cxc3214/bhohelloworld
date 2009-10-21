using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace una_register
{
    public partial class intall : Form
    {
        public intall()
        {
            InitializeComponent();
            skinEngine1.SkinFile = "Longhorn.ssk";
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (reg(checkBox1.Checked, checkBox2.Checked))
            {
                alert("修复成功");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool r = regster.doCommand("/codebase BHO_helloWorld.dll /unregister");
            if (r) { alert("卸载成功"); }
            else { alert("卸载失败"); }  
        }

        public bool reg(bool isRereg,bool isKillProcess)
        {
            try
            {
                bool isdokill = false;

                if (isRereg)
                {
                    regster.UnregisterBHO_HelloWorld();
                }

                if (isKillProcess)
                {
                    DialogResult dg = MessageBox.Show("将要停止所有IE程序，确定执行吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dg.ToString().ToLower().Equals("yes"))
                    {

                        if (regster.killProcess("iexplore") && regster.killProcess("EXPLORER") && regster.UnregisterBHO_HelloWorld())
                        {
                            isdokill = true;
                        }
                    }
                }


                if (regster.doCommand("/codebase BHO_helloWorld.dll"))
                {
                    return true;
                }

                if (isdokill == true)
                {
                    regster.startSysProcess("EXPLORER.exe");
                }
                return false;
            }
            catch (Exception e)
            {
                return false;
                throw e;
            }

            
        }


        public static void alert(string msg)
        {
            MessageBox.Show(msg, "提示",MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
