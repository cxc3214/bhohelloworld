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
            textBox1.Text = System.Environment.CurrentDirectory + "\\conf.xml";
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (reg(checkBox2.Checked))
            {
                alert("修复成功,请重新打开IE浏览器。");
            }
            else
            {
                alert("注册失败");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                RWReg rg = new RWReg("LOCAL_MACHINE");
                if (rg.GetRegVal("Software\\SimpleSoft\\BhoDir", "SysPath", "ref_null") != "ref_null")
                {
                    rg.DelRegSubKey("Software\\SimpleSoft\\BhoDir");
                }
            }

            if (regster.doCommand("/codebase BHO_helloWorld.dll /unregister"))
            {
                alert("卸载成功");
            }
            else
            {
                alert("卸载失败");
            }
        }

        public bool reg(bool isKillProcess)
        {
            try
            {
                if (textBox1.Text == "")
                {
                    alert("请选择注册文件！");
                    return false;
                }
                else
                {
                    RWReg rg = new RWReg("LOCAL_MACHINE");
                    if (rg.GetRegVal("Software\\SimpleSoft\\BhoDir", "SysPath", "ref_null") == "ref_null")
                    {
                        rg.CreateRegKey("Software\\SimpleSoft\\BhoDir");
                    }
                    rg.SetRegVal("Software\\SimpleSoft\\BhoDir", "SysPath", textBox1.Text);
                }

                if (isKillProcess)
                {
                    DialogResult dg = MessageBox.Show("将要停止所有IE程序，确定执行吗？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dg.ToString().ToLower().Equals("yes"))
                    {

                        if (regster.doCommand("/codebase BHO_helloWorld.dll") && regster.killProcess("iexplore"))
                        {
                            //regster.startSysProcess("IEXPLORE");
                            return true;
                        }
                    }
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

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog opendialog1 = new OpenFileDialog();
            opendialog1.Filter = "所有文件|*.*|配置文件|*.xml";
            opendialog1.ShowDialog();

            string path = opendialog1.FileName;
            if (path != "")
            {
                if (path.Substring(path.Length - 4, 4) != ".xml")
                {
                    alert("文件格式错误");
                }
                else
                {
                    textBox1.Text = path;
                }
            }
        }
    }
}
