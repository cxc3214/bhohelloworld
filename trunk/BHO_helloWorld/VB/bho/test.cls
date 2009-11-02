VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "simpleBHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IObjectWithSiteTLB.IObjectWithSite
Dim WithEvents IE As InternetExplorer
Attribute IE.VB_VarHelpID = -1
Dim M_Site As IUnknownVB

Private Sub IObjectWithSite_GetSite(ByVal Priid As IObjectWithSiteTLB.GUIDPtr, PPVobj As IObjectWithSiteTLB.VOIDPtr)
M_Site.QueryInterface Priid, PPVobj
End Sub

Private Sub IObjectWithSite_SetSite(ByVal pSite As IObjectWithSiteTLB.IUnknownVB)
Set M_Site = pSite
Set IE = pSite
End Sub


Private Sub IE_DownloadComplete()
    
'MsgBox url, vbInformation, "您好，IE为您服务。"

Dim elements As HTMLDocument

Set elements = IE.document

'装载elenents
intallScripts (elements)
    
End Sub

Private Sub intallScripts(document As HTMLDocument)
On Error Resume Next

Dim bhoKeyPathName As String: bhoKeyPathName = "Software\\SimpleSoft\\BhoDir"
Dim RegPathKey As String: RegPathKey = GetRegistryValue("HKEY_LOCAL_MACHINE", bhoKeyPathName, "SysPath", "null")

If RegPathKey <> "null" Then
    Dim xmlDocument As New DOMDocument
    
    xmlDocument.async = False
    
    If xmlDocument.Load(RegPathKey) = False Then
        MsgBox "配置文件载入失败!"
        Set xmlDocument = Nothing
        Exit Sub
    End If
    
    Dim webSiteNames As IXMLDOMNodeList
    Dim webSiteName As IXMLDOMNode

    
    
    Set webSiteNames = xmlDocument.selectNodes("root//website")
    '遍历配置参数
    For Each webSiteName In webSiteNames
    Dim xe As IXMLDOMElement
    Set xe = webSiteName
        
        Dim rx As New RegExp
        rx.Pattern = xe.getAttribute("url")
        rx.Global = True
        rx.IgnoreCase = False
        
        If rx.Test(document.url) Then
            '装载文件
            loadFiles webSiteName, document
            
        End If
          
    Next
End If

End Sub

Sub loadFiles(xn As IXMLDOMNode, document As HTMLDocument)
On Error Resume Next

'/*动态加载JavaScript文件
'* 需先判断是否已经加载项目；如未加载再加入该模块
'*
'*/
Dim jsFiles As IXMLDOMNodeList
Dim jsFile As IXMLDOMNode

Dim head As IHTMLDOMNode

Set jsFiles = xn.selectNodes("js")
For Each jsFile In jsFiles

Dim jsxe As IXMLDOMElement
Set jsxe = jsFile
    
    If document.getElementById(jsxe.getAttribute("keyID")) Is Empty Then
        Dim script As IHTMLElement
                           
        Set script = document.createElement("script")
        
        script.setAttribute "src", jsxe.Text
        script.setAttribute "type", "text/javascript"
        script.setAttribute "defer", jsxe.getAttribute("defer")
        script.setAttribute "id", jsxe.getAttribute("keyID")
        script.setAttribute "charset", "utf-8"
        
        Set head = document.getElementsByTagName("head").Item(0, 0)
        head.appendChild (script)
    
    End If
Next

'/*动态加载CSS文件
'* 需先判断是否已经加载项目；如未加载再加入该模块
'*
'*/

Dim cssFiles  As IXMLDOMNodeList
Dim cssFile As IXMLDOMNode

Set cssFiles = xn.selectNodes("css")
For Each cssFile In cssFiles

Dim cssxe As IXMLDOMElement
Set cssxe = jsFile

    If document.getElementById(cssxe.getAttribute("keyID")) Is Empty Then
        Dim css As IHTMLElement
        
        Set css = document.createElement("link")
        css.setAttribute "rel", "stylesheet"
        css.setAttribute "type", "text/css"
        css.setAttribute "href", cssxe.Text
        css.setAttribute "id", cssxe.getAttribute("keyID")
        Set head = document.getElementsByTagName("head").Item(0, 0)
        head.appendChild (css)
    
        'document.createStyleSheet(cssxe.InnerText, 1);//方法二
    End If
Next
End Sub




