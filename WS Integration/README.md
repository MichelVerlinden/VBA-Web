#WS Integration

In this package the framework from project Parallel-VBA-UDFs is used to create worksheet functions that compute ranges simultaneously (see Parallel-VBA-UDFs).

This allows to implement worksheet functions that use REST APIs efficiently.

To use framework, use RestWSFunction as implementation of IAsyncFun from "Parallel-VBA-UDFs" and assign object implementing IRestQuery.

### Translate Example

```VB.net
' Google Translate API
Public Function testTranslate(keyWord As String) As String
    Dim tWS As New RestWSFunction
    Dim tWeb As New TranslateQuery
    tWS.assign tWeb
    testTranslate = AsynchWSFun.asyncFun(tWS, keyWord)
End Function
```