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

Private Sub Class_Terminate()
    Set restQuery = Nothing
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

Private Function IAsyncWSFun_catchResp() As Variant
    AsynchWSFun.processResp Me
End Function

Private Function IAsyncWSFun_getName() As String
    IAsyncWSFun_getName = restQuery.Name
End Function

Private Function IAsyncWSFun_makeRequest(params As Scripting.IDictionary) As Boolean
    Debug.Print restQuery.wRequest Is Nothing
    Set restQuery.wResponse = restQuery.wClient.Execute(restQuery.wRequest)
    IAsyncWSFun_makeRequest = True
    AsynchWSFun.processResp Me
End Function

Private Function IAsyncWSFun_processResponse(strRes As String, ParamArray p() As Variant) As Boolean
    On Error GoTo errhandler
        IAsyncWSFun_processResponse = restQuery.getDataPoint(strRes, p(0))
    Exit Function
errhandler:
    IAsyncWSFun_processResponse = "#N/A"
    IAsyncWSFun_processResponse = True
End Function

Private Function IAsyncWSFun_validateRequest(params As Scripting.IDictionary) As Boolean
    Dim wRequest As WebRequest
    Set restQuery.wRequest = wRequest
    If restQuery.buildQuery(wRequest, params) Then
        IAsyncWSFun_validateRequest = True
    Else
        ' TODO
    End If
End Function

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Core processing
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Sub assign(ByRef q As IRestWSQuery)
    Set restQuery = q
End Sub


