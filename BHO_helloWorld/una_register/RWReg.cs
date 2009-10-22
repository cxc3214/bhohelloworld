using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

namespace una_register
{
    /// <summary>
    /// RWReg 的摘要说明。
    /// 注册表操作类
    /// 类库开发：吴剑冰
    /// 时间：2003年11月07日
    /// 功能：注册表操作
    /// </summary>
    public class RWReg
    {
        private static RegistryKey rootkey;
        /// <summary>
        /// 构造根键为RootKey的注册表操作类，缺省打开Current_User主键
        /// </summary>
        public RWReg(string RootKey)
        {
            switch (RootKey.ToUpper())
            {
                case "CLASSES_ROOT":
                    rootkey = Registry.ClassesRoot;
                    break;
                case "CURRENT_USER":
                    rootkey = Registry.CurrentUser;
                    break;
                case "LOCAL_MACHINE":
                    rootkey = Registry.LocalMachine;
                    break;
                case "USERS":
                    rootkey = Registry.Users;
                    break;
                case "CURRENT_CONFIG":
                    rootkey = Registry.CurrentConfig;
                    break;
                case "DYN_DATA":
                    rootkey = Registry.DynData;
                    break;
                case "PERFORMANCE_DATA":
                    rootkey = Registry.PerformanceData;
                    break;
                default:
                    rootkey = Registry.CurrentUser;
                    break;
            }
        }
        /// <summary>
        /// 读取路径为keypath，键名为keyname的注册表键值，缺省返回def
        /// </summary>
        /// <param name="keypath"></param>
        /// <param name="keyname"></param>
        /// <param name="def"></param>
        /// <returns></returns>
        public string GetRegVal(string keypath, string keyname, string def)
        {
            try
            {
                RegistryKey key = rootkey.OpenSubKey(keypath);
                return key.GetValue(keyname, (object)def).ToString();
            }
            catch
            {
                return def;
            }
        }
        /// <summary>
        /// 设置路径为keypath，键名为keyname的注册表键值为keyval
        /// </summary>
        /// <param name="keypath"></param>
        /// <param name="keyname"></param>
        /// <param name="keyval"></param>
        public bool SetRegVal(string keypath, string keyname, string keyval)
        {
            try
            {
                RegistryKey key = rootkey.OpenSubKey(keypath, true);
                key.SetValue(keyname, (object)keyval);
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 创建路径为keypath的键
        /// </summary>
        /// <param name="keypath"></param>
        /// <returns></returns>
        public RegistryKey CreateRegKey(string keypath)
        {
            try
            {
                return rootkey.CreateSubKey(keypath);
            }
            catch
            {
                return null;
            }
        }
        /// <summary>
        /// 删除路径为keypath的子项
        /// </summary>
        /// <param name="keypath"></param>
        /// <returns></returns>
        public bool DelRegSubKey(string keypath)
        {
            try
            {
                rootkey.DeleteSubKey(keypath);
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 删除路径为keypath的子项及其附属子项
        /// </summary>
        /// <param name="keypath"></param>
        /// <returns></returns>
        public bool DelRegSubKeyTree(string keypath)
        {
            try
            {
                rootkey.DeleteSubKeyTree(keypath);
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 删除路径为keypath下键名为keyname的键值
        /// </summary>
        /// <param name="keypath"></param>
        /// <param name="keyname"></param>
        /// <returns></returns>
        public bool DelRegKeyVal(string keypath, string keyname)
        {
            try
            {
                RegistryKey key = rootkey.OpenSubKey(keypath, true);
                key.DeleteValue(keyname);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
