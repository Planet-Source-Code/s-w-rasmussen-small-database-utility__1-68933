VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VicsDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Const S_OK As Long = &H0

Public Event VicDLStart()                                                                       'notify when current download is started
Public Event VicDLDone()                                                                        'notify when current download is finished
Public Event VicDLCrash(ByVal VicErrNum As Long, VicErrDescr As String, Cancelled As Boolean)   'notify when download was interrupted
Public Event VicDLCancelled()                                                                   'notify when a cancel request was received
Public Event VicDLProg(ByVal VicBytesIn As Long, ByVal VicTotalBytes As Long)                   'notify of progress during download

Implements olelib.IBindStatusCallback                                                           'initialize the IBindStatusCallback interface...

Private VicBinding As IBinding                                                                  ' set-up so I can access as I wish

Private Sub Class_Terminate()
    
    If VicBinding Is Nothing Then
        Exit Sub
    Else
        Set VicBinding = Nothing
    End If
    
End Sub

Private Sub IBindStatusCallback_OnProgress(ByVal ulProgress As Long, ByVal ulProgressMax As Long, ByVal ulStatusCode As olelib.BINDSTATUS, ByVal szStatusText As Long)
    
    If ulProgressMax > 0 Then
        RaiseEvent VicDLProg(ulProgress, ulProgressMax) 'report progress to the event
    End If
    DoEvents
    
End Sub

Private Sub IBindStatusCallback_OnStartBinding(ByVal dwReserved As Long, ByVal pib As olelib.IBinding)
    
    Set VicBinding = pib
    RaiseEvent VicDLStart

End Sub

Private Sub IBindStatusCallback_OnStopBinding(ByVal hresult As Long, ByVal szError As Long)
    
    If hresult = 1 Or hresult = S_OK Then
        RaiseEvent VicDLDone
    Else
        Dim Cancelled As Boolean
        If VicBinding Is Nothing Then
            Cancelled = True
        Else
            Cancelled = False
        End If
        RaiseEvent VicDLCrash(hresult, GetVicError(hresult), Cancelled) 'report the failure to the event for handling
    End If
    Set VicBinding = Nothing
    
End Sub

Private Sub IBindStatusCallback_GetBindInfo(grfBINDF As olelib.BINDF, pbindinfo As olelib.BINDINFO)

End Sub

Private Function IBindStatusCallback_GetPriority() As Long

End Function

Private Sub IBindStatusCallback_OnDataAvailable(ByVal grfBSCF As olelib.BSCF, ByVal dwSize As Long, pformatetc As olelib.FORMATETC, pStgmed As olelib.STGMEDIUM)

End Sub

Private Sub IBindStatusCallback_OnLowResource(ByVal reserved As Long)

End Sub

Private Sub IBindStatusCallback_OnObjectAvailable(riid As olelib.UUID, ByVal pUnk As stdole.IUnknown)

End Sub

Public Function DownloadSuccess(ByVal File2DownLoad As String, ByVal File2Save As String)
    
Dim DownLoadResult As Long

    If VicBinding Is Nothing Then
        DownLoadResult = olelib.URLDownloadToFile(Nothing, File2DownLoad, File2Save, 0, Me)
        DownloadSuccess = (DownLoadResult = olelib.S_OK)
    End If
    
End Function

Public Sub KillVic()

On Error GoTo NoBind2
    
    RaiseEvent VicDLCancelled
    Set VicBinding = Nothing
    
    Exit Sub

NoBind2:
    Resume Next
End Sub

Private Function GetVicError(ByVal VicErrNum As Long) As String

    Select Case VicErrNum
        Case INET_E_AUTHENTICATION_REQUIRED
            GetVicError = "Authentication Failure."
        Case INET_E_CANNOT_CONNECT
            GetVicError = "Cannot Connect"
        Case INET_E_CANNOT_INSTANTIATE_OBJECT
            GetVicError = "Cannot Instantiate Object."
        Case INET_E_CANNOT_LOAD_DATA
            GetVicError = "Cannot Load Data."
        Case INET_E_CANNOT_LOCK_REQUEST
            GetVicError = "Cannot Lock Request."
        Case INET_E_CANNOT_REPLACE_SFP_FILE
            GetVicError = "Cannot Replace SFP File."
        Case INET_E_CODE_DOWNLOAD_DECLINED
            GetVicError = "Code Download Declined."
        Case INET_E_CONNECTION_TIMEOUT
            GetVicError = "Connection Timeout."
        Case INET_E_DATA_NOT_AVAILABLE
            GetVicError = "Data Not Available."
        Case INET_E_DEFAULT_ACTION
            GetVicError = "Default Action."
        Case INET_E_DOWNLOAD_FAILURE
            GetVicError = "Download Failure."
        Case INET_E_INVALID_REQUEST
            GetVicError = "Invalid Request."
        Case INET_E_INVALID_URL
            GetVicError = "Invalid URL."
        Case INET_E_NO_SESSION
            GetVicError = "No Session."
        Case INET_E_NO_VALID_MEDIA
            GetVicError = "No Valid Media."
        Case INET_E_OBJECT_NOT_FOUND
            GetVicError = "File Not Found."
        Case INET_E_QUERYOPTION_UNKNOWN
            GetVicError = "QueryOption Unknown."
        Case INET_E_REDIRECT_FAILED
            GetVicError = "Redirect Failed."
        Case INET_E_REDIRECT_TO_DIR
            GetVicError = "Redirect To Dir."
        Case INET_E_REDIRECTING
            GetVicError = "Redirecting."
        Case INET_E_RESOURCE_NOT_FOUND
            GetVicError = "Resource Not Found."
        Case INET_E_RESULT_DISPATCHED
            GetVicError = "Result Dispatched."
        Case INET_E_SECURITY_PROBLEM
            GetVicError = "Security Problem."
        Case INET_E_UNKNOWN_PROTOCOL
            GetVicError = "Unknown Protocol."
        Case Else
            GetVicError = "Unknown."
   End Select
   
End Function

Public Function DeleteVicCache(Filename As String) As Long
    
    DeleteVicCache = DeleteUrlCacheEntry(Filename)
    
End Function
