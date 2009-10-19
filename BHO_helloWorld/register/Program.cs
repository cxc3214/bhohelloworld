using System;
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
            string bhoKeyPathName = "Software\\SimpleSoft\\BhoDir";
            string path = Registry.LocalMachine.OpenSubKey(bhoKeyPathName, true).GetValue("SysPath", "null").ToString();

            System.Windows.Forms.MessageBox.Show(path);

        }


    }
}
