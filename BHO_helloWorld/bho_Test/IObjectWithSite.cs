using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace BrowserMonitor
{

    #region 引入IObjectWithSite接口
    [ComImport(), Guid("fc4801a3-2ba9-11cf-a229-00aa003d7352")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IObjectWithSite
    {
        void SetSite([In, MarshalAs(UnmanagedType.IUnknown)] object site);
        void GetSite(ref Guid guid, [MarshalAs(UnmanagedType.IUnknown)] out object site);
    }
    #endregion 

    #region 定义默认BHO COM接口
    [GuidAttribute("181C179B-7CC9-4457-8C1D-4B45E7C8589D")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsDual)]
    public interface IObserver
    {

    }
    #endregion 


}
