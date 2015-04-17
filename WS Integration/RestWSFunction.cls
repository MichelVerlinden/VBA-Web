VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RestWSFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IAsyncWSFun

Private idReq As Integer

Private restQuery As IRestWSQuery

Public web_Client As WebClient
Public web_Request As WebRequest
Public web_Response As WebResponse

''
' @property Http
' @type WebHttpRequest
''
Public WithEvents Http As WinHttpRequest
Attribute Http.VB_VarHelpID = -1

Private Sub Class_Terminate()
    Set restQuery = Nothing
    Set Http = Nothing
    Set web_Request = Nothing
    Set web_Response = Nothing
    Set web_Client = Nothing
End Sub

Public Sub assign(ByRef q As IRestWSQuery)
    Set restQuery = q
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Implementation of IAsyncWSFun
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Private Property Let IAsyncWSFun_Id(ByVal RHS As Integer)
    idReq = RHS
End Property

Private Property Get IAsyncWSFun_Id() As Integer
    IAsyncWSFun_Id = idReq
End Property

Public Sub IAsyncWSFun_catchResp()
    AsynchWSFun.processResp Me
End Sub

Private Function IAsyncWSFun_getName() As String
    IAsyncWSFun_getName = restQuery.Name
End Function

Private Function IAsyncWSFun_makeRequest(params As Scripting.IDictionary) As Boolean
    If restQuery.initClient(web_Client) Then
        PrepareAndExecuteRequest web_Request
    End If
    IAsyncWSFun_makeRequest = True
End Function

Private Function IAsyncWSFun_processResponse(strRes As String, ParamArray p() As Variant) As Boolean
    On Error GoTo errhandler
        IAsyncWSFun_processResponse = restQuery.getDataPoint(strRes, web_Response, p(0))
    Exit Function
errhandler:
    IAsyncWSFun_processResponse = "#N/A"
    IAsyncWSFun_processResponse = True
End Function

Private Function IAsyncWSFun_validateRequest(params As Scripting.IDictionary) As Boolean
    If restQuery.buildRequest(web_Request, params) Then
        IAsyncWSFun_validateRequest = True
    Else
        ' TODO
    End If
End Function

' ============================================= '
' Asynchonous Execution
' ============================================= '

''
' Once everything has been prepared, execute request
'
' @internal
' @method PrepareAndExecuteRequest
' @param {RestRequest} Request
''
Private Sub PrepareAndExecuteRequest(Request As WebRequest)
    On Error GoTo web_ErrorHandling

    Set Me.Http = web_Client.PrepareHttpRequest(Request)

    Me.Http.Send Request.Body
    Exit Sub
    
web_ErrorHandling:
    
    Set Me.Http = Nothing
    Set Request = Nothing
    
    ' Rethrow error
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

''
' Handle timeout
'
' @internal
' @method TimedOut
''
Private Sub TimedOut()
    Dim web_Response As New WebResponse
    
    web_StopTimeoutTimer
    WebHelpers.LogDebug "Timed out", "WebAsyncWrapper.TimedOut"
    
    ' Callback
    web_Response.StatusCode = WebStatusCode.RequestTimeout
    web_Response.StatusDescription = "Request Timeout"
End Sub

' Start timeout timer
Private Sub web_StartTimeoutTimer()
    Dim web_TimeoutS As Long

    If WebHelpers.AsyncRequests Is Nothing Then: Set WebHelpers.AsyncRequests = New Dictionary
    
    ' Round ms to seconds with minimum of 1 second if ms > 0
    web_TimeoutS = Round(web_Client.TimeoutMs / 1000, 0)
    If web_Client.TimeoutMs > 0 And web_TimeoutS = 0 Then
        web_TimeoutS = 1
    End If
    
    WebHelpers.AsyncRequests.Add web_Request.Id, Me
    Application.OnTime Now + TimeValue("00:00:" & web_TimeoutS), "'WebHelpers.OnTimeoutTimerExpired """ & web_Request.Id & """'"
End Sub

' Stop timeout timer
Private Sub web_StopTimeoutTimer()
    If Not WebHelpers.AsyncRequests Is Nothing And Not web_Request Is Nothing Then
        If WebHelpers.AsyncRequests.Exists(web_Request.Id) Then
            WebHelpers.AsyncRequests.Remove web_Request.Id
        End If
    End If
End Sub

' Process asynchronous requests
Private Sub Http_OnResponseFinished()
    Set web_Response = New WebResponse
    
  '  web_StopTimeoutTimer
    
    ' Callback
    web_Response.CreateFromHttp web_Client, web_Request, Me.Http
    
    Me.IAsyncWSFun_catchResp
End Sub

