using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Text;
using SHDocVw;
using mshtml;
using System.Xml;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace BHO_HelloWorld
{

    [
    ComVisible(true),
    Guid("8a194578-81ea-4850-9911-13ba2d71efbd"),
    ClassInterface(ClassInterfaceType.None)
    ]
    public class BHO : IObjectWithSite
    {

        WebBrowser webBrowser;
        HTMLDocument document;

        public void OnDocumentComplete(object pDisp, ref object URL)
        {
            //alert("OnDocumentComplete"+ URL.ToString());
            document = (HTMLDocument)webBrowser.Document;
            intallScripts(document);
        }

        public void DownloadComplete()
        {
            //alert("DownloadComplete");
            document = (HTMLDocument)webBrowser.Document;
            intallScripts(document);
        }

        public void DownloadBegin()
        {
            alert("DownloadBegin");
            document = (HTMLDocument)webBrowser.Document;
            intallScripts(document);
        }
        //public void OnBeforeNavigate2(object pDisp, ref object URL, ref object Flags, ref object TargetFrameName, ref object PostData, ref object Headers, ref bool Cancel)
        public void NavigateComplete2(object pDisp, ref object URL)
        {
            //alert("NavigateComplete2"+ URL.ToString());
            document = (HTMLDocument)webBrowser.Document;
            intallScripts(document);
        }

        #region BHO Internaled Functions
        public bool intallScripts(HTMLDocument document)
        {
            string bhoKeyPathName = "Software\\SimpleSoft\\BhoDir";
            RegistryKey RegPathKey = Registry.LocalMachine.OpenSubKey(bhoKeyPathName, true);
            if (RegPathKey != null)
            {
                string path = RegPathKey.GetValue("SysPath", "null").ToString();
                if (Directory.Exists(path))
                {
                    XmlDocument doc = readXml(path + "\\conf.xml");
                    XmlNodeList webSiteNames = doc.SelectNodes("root/website");
                    foreach (XmlNode webSiteName in webSiteNames)
                    {
                        XmlElement xe = (XmlElement)webSiteName;
                        Regex rx = new Regex(@""+ xe.GetAttribute("url")+"");

                        if (rx.IsMatch(document.url))
                        {
                            loadFiles(webSiteName,document);
                        }
                    }
                }
                else
                {
                    alert("注册文件丢失，请重新安装插件。");
                    return false;
                }
            }
            else
            {
                alert("注册文件丢失，请重新安装插件。");
                return false;
            }
            return true;
        }

        public void loadFiles(XmlNode xn,HTMLDocument document)
        {
            try
            {

                /*动态加载JavaScript文件
                * 需先判断是否已经加载项目；如未加载再加入该模块
                * 
                */

                XmlNodeList jsFiles = xn.SelectNodes("js");
                foreach (XmlNode jsFile in jsFiles)
                {
                    XmlElement jsxe = (XmlElement)jsFile;

                    if (document.getElementById(jsxe.GetAttribute("keyID")) == null)
                    {                        
                        IHTMLElement script = document.createElement("script");
                        script.setAttribute("src", jsxe.InnerText, 0);
                        script.setAttribute("type", "text/javascript", 0);
                        script.setAttribute("defer", jsxe.GetAttribute("defer"), 0);
                        script.setAttribute("id", jsxe.GetAttribute("keyID"), 0); 
                        script.setAttribute("charset", "utf-8", 0);
                        IHTMLDOMNode head = (IHTMLDOMNode)document.getElementsByTagName("head").item(0, 0);
                        head.appendChild((IHTMLDOMNode)script);

                    }
                }

                /*动态加载CSS文件
                 * 需先判断是否已经加载项目；如未加载再加入该模块
                 * 
                 */

                XmlNodeList cssFiles = xn.SelectNodes("css");
                foreach (XmlNode cssFile in cssFiles)
                {
                    XmlElement cssxe = (XmlElement)cssFile;

                    if (document.getElementById(cssxe.GetAttribute("keyID")) == null)
                    {
                        
                        IHTMLElement css = document.createElement("link");
                        css.setAttribute("rel", "stylesheet", 0);
                        css.setAttribute("type", "text/css", 0);
                        css.setAttribute("href", cssxe.InnerText, 0);
                        css.setAttribute("id", cssxe.GetAttribute("keyID"), 0);
                        IHTMLDOMNode head = (IHTMLDOMNode)document.getElementsByTagName("head").item(0, 0);
                        head.appendChild((IHTMLDOMNode)css);

                        //document.createStyleSheet(cssxe.InnerText, 1);//方法二
                    }
                }
            }
            catch (Exception e)
            {
                //alert(e.Message);
                throw e;
            }

        }

        public bool IsType(object obj, string type)
        {
            Type t = Type.GetType(type, true, true);
            
            alert(t.ToString());
            return t == obj.GetType() || obj.GetType().IsSubclassOf(t);
        }



        public XmlDocument readXml(string xmlPath)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlPath);
            return doc;
        }

        public void alert(string msg) 
        {
            System.Windows.Forms.MessageBox.Show(msg, "提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning); 
        }

        #endregion

        #region BHO Internal Functions
        public static string BHOKEYNAME = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
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

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHOKEYNAME, true);
            string guid = type.GUID.ToString("B");

            if (registryKey != null)
                registryKey.DeleteSubKey(guid, false);
        }

        public int SetSite(object site)
        {

            if (site != null)
            {
                webBrowser = (WebBrowser)site;
                
                webBrowser.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                //webBrowser.NavigateComplete2 += new DWebBrowserEvents2_NavigateComplete2EventHandler(this.NavigateComplete2);
                webBrowser.DownloadComplete += new DWebBrowserEvents2_DownloadCompleteEventHandler(this.DownloadComplete);
               // webBrowser.DownloadBegin += new DWebBrowserEvents2_DownloadBeginEventHandler(this.DownloadBegin);
            }
            else
            {
                webBrowser.DocumentComplete -= new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
               // webBrowser.NavigateComplete2 -= new DWebBrowserEvents2_NavigateComplete2EventHandler(this.NavigateComplete2);

                webBrowser.DownloadComplete -= new DWebBrowserEvents2_DownloadCompleteEventHandler(this.DownloadComplete);
               // webBrowser.DownloadBegin -= new DWebBrowserEvents2_DownloadBeginEventHandler(this.DownloadBegin);
                webBrowser = null;
            }

            return 0;

        }

        public int GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(webBrowser);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
            Marshal.Release(punk);

            return hr;
        }

        #endregion


    }
}
