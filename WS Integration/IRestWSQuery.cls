VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRestWSQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String

Public Function initClient(ByRef Client As WebClient) As Boolean
End Function

Public Function buildRequest(ByRef Request As WebRequest, ByRef params As Dictionary) As Boolean
End Function

Public Function getDataPoint(ByRef strRes As String, ByRef Response As WebResponse, ParamArray p() As Variant) As Boolean
End Function
