Attribute VB_Name = "WSFunctions"
'/*
'' WSFunctions
'' Module where worksheet functions should be defined
''
'' In order to define "parralel" function: Implement IAsynchWSFun
'' in a class module and use the object in AsychWSFun.asyncFun(<your object>, <your parameters>)
''
''
'' author : Michel Verlinden
'' 17/03/2014
''
'' TODO :   Add generic argument validator
''          Function registration
''
'*/
Option Explicit
'Option Private Module ' comment this if not registering functions

' Twitter Sentiment
Public Function testTwitter(keyWord As String) As String
    Dim tWS As New RestWSFunction
    Dim tWeb As New TwitterQuery
    tWS.assign tWeb
    testTwitter = AsynchWSFun.asyncFun(tWS, keyWord)
End Function
