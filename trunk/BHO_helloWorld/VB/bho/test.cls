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
    
    MsgBox url, vbInformation, "您好，IE为您服务。"
    
    Dim elements As HTMLDocument
    
    Set elements = IE.document
    
    
End Sub

