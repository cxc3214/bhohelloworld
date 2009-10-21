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
        public bool doCommand(string commStrings )
        {
            try
            {
                ProcessStartInfo start = new ProcessStartInfo("regasm.exe");
                start.Arguments = commStrings;//设置命令参数
                start.RedirectStandardOutput = true;//
                start.RedirectStandardInput = true;//
                start.CreateNoWindow = true;
                start.UseShellExecute = false;//是否指定操作系统外壳进程启动程序
                Process p = Process.Start(start);
                /*
                StreamReader reader = p.StandardOutput;//截取输出流
                string line = reader.ReadLine();//每次读取一行
                while (!reader.EndOfStream)
                {
                    Console.WriteLine(line+" ");
                    line = reader.ReadLine();
                }
                 */
                return true;
                p.WaitForExit();//等待程序执行完退出进程
                p.Close();//关闭进程
            }
            catch (Exception e)
            {
                return false;
                throw e;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool r = doCommand("BHO_helloWorld.dll");
            if (r) { alert("修复成功"); }
            else { alert("修复失败"); }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool r = doCommand("BHO_helloWorld.dll /unregister");
            if (r) { alert("卸载成功"); }
            else { alert("卸载失败"); }
            
        }

        public static void alert(string msg)
        {
            MessageBox.Show(msg, "提示",MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
