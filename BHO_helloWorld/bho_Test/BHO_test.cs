using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Text;
using SHDocVw;
using System.Xml;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace BrowserMonitor
{

    /// <summary> 
    ///定义BHO类,此类由浏览器实例化 
    /// </summary> 
    /// 

    [
    ComVisible(true),
    Guid("B29E305D-BC4D-4a80-B522-B0ABC9EBDFFC"),
    ClassInterfaceAttribute(ClassInterfaceType.None)
    ]
    //由CreateGUID程序生成 
    [ProgIdAttribute("Observer.BrowserMonitor")]
    public class BrowserMonitor : IObserver, IObjectWithSite
    {

        
            
        IWebBrowser2 browser;  //浏览器对象 
        DWebBrowserEvents2_Event browserEvents; //浏览器事件 

        #region IObjectWithSite 成员
        /// <summary> 
        /// IE调用此方法,并传递指向容器Site的IUnknown指针,由此我们可以获得IWebBrowser2接口 
        /// 并挂载DWebBrowserEvents2事件 
        /// </summary> 
        /// <param name="site">容器Site的IUnknown指针</param>
        /// 

        
        public void SetSite(object site)
        {
            // TODO:  添加 BrowserMonitor.SetSite 实现 
            #region 取得 IWebBrowser2 引用
            if (browser != null)
                Release();
            if (site == null)
                return;
            browser = site as IWebBrowser2;
            #endregion
            
            #region 检查名称,当且仅当为IEXPLORE.EXE时加载
            string hostName = browser.FullName;
            if (!(hostName.ToUpper().EndsWith("IEXPLORE.EXE")))
            {
                Release();
                return;
            }
            #endregion
            #region 挂载浏览器事件
            browserEvents = browser as DWebBrowserEvents2_Event;
            if (browserEvents != null)
            {
                browserEvents.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
            }
            else
            {
                Release();
                return;
            }
            #endregion
        }
        /// <summary> 
        /// 调用者调用此方法以获得前面浏览器发送给SetSite()方法的浏览器对象 
        /// </summary> 
        /// <param name="guid">请求Site接口对象的GUID</param> 
        /// <param name="site">返回的Site接口对象</param> 
        public void GetSite(ref System.Guid guid, out object site)
        {
            // TODO:  添加 BrowserMonitor.GetSite 实现 
            site = null;
            if (browser != null)
            {
                IntPtr pSite = IntPtr.Zero;
                IntPtr pUnk = Marshal.GetIUnknownForObject(browser); //引用计数增加 
                Marshal.QueryInterface(pUnk, ref guid, out pSite); //引用计数增加 
                Marshal.Release(pUnk);  //引用计数减少 
                Marshal.Release(pUnk); //引用计数减少 
                if (!pSite.Equals(IntPtr.Zero))
                {
                    site = pSite;
                }
                else
                {
                    // 若找不到请求的接口，将返回E_NOINTERFACE 
                    Release();
                    Marshal.ThrowExceptionForHR(E_NOINTERFACE);
                }
            }
            else
            {
                // 若找不到请求的接口对象，将返回E_FAIL 
                Release();
                Marshal.ThrowExceptionForHR(E_FAIL);
            }
        }
        #endregion


        #region 用于挂载的自定义函数
        protected void OnDocumentComplete(object display, ref object url)
        {
            try
            {
                if (Marshal.Equals(browser, display))
                {
                    StreamWriter sw = new StreamWriter(@"C:\Test.txt", true);
                    sw.WriteLine(url.ToString());
                    sw.Close();
                }
            }
            catch
            {
                Release();
                Marshal.ThrowExceptionForHR(E_FAIL);
            }
        }
        #endregion


        #region 定义HRESULT值 : 预定义的COMException值
        //在未检查的上下文中，如果表达式产生目标类型范围之外的值，则结果被截断 
        const int E_FAIL = unchecked((int)0x80004005); //失败 
        const int E_NOINTERFACE = unchecked((int)0x80004002);//QueryInterface时,接口不存在 
        //Marshal.ThrowExceptionForHR E_FAIL;
        //Marshal.ThrowExceptionForHR E_NOINTERFACE;
        //Marshal.ThrowExceptionForHR (E_FAIL); 
        //Marshal.ThrowExceptionForHR(E_NOINTERFACE); 

        #endregion
        //使用一下方式，释放对象： 
        #region Release操作
        protected void Release()
        {
            if (browserEvents != null)
            {
                Marshal.ReleaseComObject(browserEvents);
                browserEvents = null;
            }
            if (browser != null)
            {
                Marshal.ReleaseComObject(browser);
                browser = null;
            }
        }
        #endregion



        #region 挂载注册/注销操作
        /// <summary> 
        /// 在注册COM组件时,由运行时调用 
        /// </summary> 
        [ComRegisterFunction]
        public static void Register(Type type)
        {
            // 注册BHO组件 

            string BHOKEYNAME = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects";

            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHOKEYNAME, true);

            if (registryKey == null)
                registryKey = Registry.LocalMachine.CreateSubKey(BHOKEYNAME);

            string guid = type.GUID.ToString("B");
            RegistryKey ourKey = registryKey.OpenSubKey(guid);

            if (ourKey == null)
                ourKey = registryKey.CreateSubKey(guid);

            ourKey.SetValue("Alright", 1);
            registryKey.Close();
            ourKey.Close();

        }
        /// <summary> 
        /// 在注销COM组件时,由运行时调用 
        /// </summary> 
        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            // 注销BHO组件 
            string guid = type.GUID.ToString("B");
            RegistryKey rkey =
         Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects");
            rkey.DeleteSubKey(guid, false);
        }

        #endregion
        


    }
}

    
    
 
