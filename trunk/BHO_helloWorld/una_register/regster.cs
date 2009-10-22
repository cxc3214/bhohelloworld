using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;
using Microsoft.Win32;

namespace una_register
{
    class regster
    {
        public regster()
        {

        }
        public static bool doCommand(string commStrings)
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


        public static bool killProcess(string name)
        {

            try
            {
                foreach (Process thisproc in Process.GetProcesses())
                {
                    Console.WriteLine(thisproc.ProcessName);
                    if (thisproc.ProcessName.ToLower().Equals(name))
                    {
                        thisproc.Kill();
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                return false;
                throw e;
            }
        }

        public static bool startSysProcess(string name)
        {

            try
            {
                Process p = new Process();
                p.StartInfo.FileName = name;
                p.Start();

                return true;
            }
            catch (Exception e)
            {
                return false;
                throw e;
            }
        }
    }
}
