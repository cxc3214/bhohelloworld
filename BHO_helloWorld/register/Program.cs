using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.IO;

namespace register
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                ProcessStartInfo start = new ProcessStartInfo("regasm.exe");
                start.Arguments = "BHO_helloWorld.dll /unregister";//设置命令参数
                start.RedirectStandardOutput = true;//
                start.RedirectStandardInput = true;//
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

                p.WaitForExit();//等待程序执行完退出进程
                p.Close();//关闭进程
            }
            catch (Exception e)
            {
                alert(e.Message);
            }
        }

        public static void alert(string msg)
        {
            System.Windows.Forms.MessageBox.Show(msg, "提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
        }



    }
}
