VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TranslateQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*
'' TranslateQuery
'' Translate from English to German on the worksheet
''
''
'' author : Michel Verlinden
'' 19/04/2014
''
''
'*/


Option Explicit
Implements IRestWSQuery

Private toTranslate As String
Private translated As String


Private Function IRestWSQuery_buildRequest(Request As WebRequest, params As Scripting.IDictionary) As Boolean
    If Request Is Nothing Then
        Set Request = New WebRequest
    End If
    Dim strT As String
    strT = makeText(params)
    Request.AddQuerystringParam "key", Credentials.Values("Google")("api_key")
    Request.AddQuerystringParam "target", "de"
    Request.AddQuerystringParam "q", strT
    toTranslate = Right(strT, Len(strT) - 3)
    IRestWSQuery_buildRequest = True
End Function

Private Function IRestWSQuery_getDataPoint(strRes As String, Response As WebResponse, ParamArray p() As Variant) As Boolean
    If Response Is Nothing Then
        Set Response = New WebResponse
    End If
    If Response.StatusCode = WebStatusCode.Ok Then
        If Len(translated) = 0 Then
            translated = Response.Data("data")("translations")(1)("translatedText")
            translated = Right(translated, Len(translated) - 3)
        End If
        Dim i As Integer, sIter As Variant, resArray() As String
        For Each sIter In Split(toTranslate, "</p><p>")
            If StrComp(VBA.CStr(Join(p(0))), VBA.CStr(sIter)) = 0 Then
                resArray = Split(translated, "</p><p>")
                strRes = resArray(i)
            End If
            i = i + 1
        Next sIter
    Else
        Debug.Print Response.Content
    End If
    IRestWSQuery_getDataPoint = True
End Function

Private Function IRestWSQuery_initClient(Client As WebClient) As Boolean
    If Client Is Nothing Then
        Set Client = New WebClient
        Client.BaseUrl = "https://www.googleapis.com/language/translate/v2"
        IRestWSQuery_initClient = True
    Else
        IRestWSQuery_initClient = False
    End If
End Function

Private Property Let IRestWSQuery_Name(ByVal RHS As String)
End Property

Private Property Get IRestWSQuery_Name() As String
    IRestWSQuery_Name = "Translate"
End Property

Private Function makeText(params As Dictionary) As String
    makeText = "<p>"
    Dim p As Variant
    For Each p In params.Items
        makeText = makeText & p(0) & "</p><p>"
    Next p
    makeText = Left(makeText, Len(makeText) - 7)
End Function
