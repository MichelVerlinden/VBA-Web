VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TwitterQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IRestWSQuery

Private pTwitterClient As WebClient
Private pTwitterRequest As WebRequest
Private pTwitterResponse As WebResponse
Private pTwitterKey As String
Private pTwitterSecret As String

Private Property Let IRestWSQuery_Name(ByVal RHS As String)
End Property

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Properties
'  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Property Get IRestWSQuery_Name() As String
    IRestWSQuery_Name = "TwitterSentiment"
End Property

Private Property Set IRestWSQuery_wClient(ByVal RHS As WebClient)
    Set pTwitterClient = RHS
End Property

Private Property Get IRestWSQuery_wClient() As WebClient
    If pTwitterClient Is Nothing Then

        Set pTwitterClient = New WebClient
        pTwitterClient.BaseUrl = "https://api.twitter.com/1.1/"
        
        Dim Auth As New TwitterAuthenticator
        Auth.Setup _
            ConsumerKey:=TwitterKey, _
            ConsumerSecret:=TwitterSecret
        Set pTwitterClient.Authenticator = Auth
    End If
    Set IRestWSQuery_wClient = pTwitterClient
End Property

' Implement caching for Consumer Key, Consumer Secret, and WebClient
Private Property Get TwitterKey() As String
    If pTwitterKey = "" Then
        If Credentials.Loaded Then
            pTwitterKey = Credentials.Values("Twitter")("key")
        Else
            pTwitterKey = InputBox("Please Enter Twitter Consumer Key")
        End If
    End If
    TwitterKey = pTwitterKey
End Property

Private Property Get TwitterSecret() As String
    If pTwitterSecret = "" Then
        If Credentials.Loaded Then
            pTwitterSecret = Credentials.Values("Twitter")("secret")
        Else
            pTwitterSecret = InputBox("Please Enter Twitter Consumer Secret")
        End If
    End If
    TwitterSecret = pTwitterSecret
End Property

Private Function IRestWSQuery_buildQuery(ByRef wRequest As WebRequest, params As Scripting.IDictionary) As Boolean
    Set pTwitterRequest = New WebRequest
    pTwitterRequest.Resource = "search/tweets.{format}"
    pTwitterRequest.Format = json
    pTwitterRequest.AddUrlSegment "format", "json"
    pTwitterRequest.AddQuerystringParam "q", makeTwitterQuery(params)
    pTwitterRequest.AddQuerystringParam "lang", "en"
    pTwitterRequest.AddQuerystringParam "count", 20
    pTwitterRequest.Method = HttpGet
    IRestWSQuery_buildQuery = True
End Function

Private Function IRestWSQuery_getDataPoint(strRes As String, ParamArray p() As Variant) As Boolean
    IRestWSQuery_getDataPoint = True
    Debug.Print pTwitterResponse.Data(0)
    strRes = "Completed"
End Function

Private Property Set IRestWSQuery_wRequest(ByVal RHS As WebRequest)
    Set pTwitterRequest = RHS
End Property
    
Private Property Get IRestWSQuery_wRequest() As WebRequest
    Set IRestWSQuery_wRequest = pTwitterRequest
End Property

Private Property Set IRestWSQuery_wResponse(ByVal RHS As WebResponse)
    Set pTwitterResponse = RHS
End Property

Private Property Get IRestWSQuery_wResponse() As WebResponse
    Set IRestWSQuery_wResponse = pTwitterResponse
End Property

Private Function makeTwitterQuery(params As Dictionary) As String
    makeTwitterQuery = "Excel"
End Function
