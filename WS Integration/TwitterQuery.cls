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

Private pTwitterKey As String
Private pTwitterSecret As String

Private Function IRestWSQuery_buildRequest(Request As WebRequest, params As Scripting.IDictionary) As Boolean
    Set Request = New WebRequest
    Request.Resource = "search/tweets.{format}"
    Request.Format = json
    Request.AddUrlSegment "format", "json"
    Request.AddQuerystringParam "q", makeTwitterQuery(params)
    Request.AddQuerystringParam "lang", "en"
    Request.AddQuerystringParam "count", 20
    Request.Method = HttpGet
    IRestWSQuery_buildRequest = True
End Function

Private Function IRestWSQuery_getDataPoint(strRes As String, Response As WebResponse, ParamArray p() As Variant) As Boolean
    Debug.Print Response.Content
    strRes = "Completed"
    IRestWSQuery_getDataPoint = True
End Function

Private Function IRestWSQuery_initClient(Client As WebClient) As Boolean
    If Client Is Nothing Then

        Set Client = New WebClient
        Client.BaseUrl = "https://api.twitter.com/1.1/"
        
        Dim Auth As New TwitterAuthenticator
        Auth.Setup _
            ConsumerKey:=TwitterKey, _
            ConsumerSecret:=TwitterSecret
        Set Client.Authenticator = Auth
    End If
    IRestWSQuery_initClient = True
End Function

Private Property Let IRestWSQuery_Name(ByVal RHS As String)
End Property

Private Property Get IRestWSQuery_Name() As String
    IRestWSQuery_Name = "Twitter"
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

Private Function makeTwitterQuery(params As Dictionary) As String
    Dim p As Variant
    For Each p In params.Items
        makeTwitterQuery = makeTwitterQuery & p(0) & "+OR+"
    Next p
    makeTwitterQuery = Left(makeTwitterQuery, Len(makeTwitterQuery) - 4)
End Function
