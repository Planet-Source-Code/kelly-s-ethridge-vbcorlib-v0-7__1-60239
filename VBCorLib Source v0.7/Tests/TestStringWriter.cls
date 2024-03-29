VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStringWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   TestStringWriter
'
Option Explicit

Private sw As StringWriter

Public Sub Setup()
    Dim c As CultureInfo
    
    ' the tests expect the formatting to be in usa formats.
    Set c = newcultureinfo("en-US")
    Set sw = NewStringWriter(provider:=c)
End Sub

Public Sub TestWriteLongs()
    Dim i As Long
    Dim s As String
    
    For i = 1 To 10
        sw.WriteValue i
    Next i
    
    Assert.Equals "12345678910", sw.ToString
End Sub

Public Sub TestWriteDouble()
    sw.WriteValue 4313234.22347723
    Assert.Equals "4313234.22347723", sw.ToString
End Sub

Public Sub TestWriteDate()
    Dim d As Date
    d = #2/2/2004 3:00:00 PM#
    
    sw.WriteValue d
    Assert.Equals "2/2/2004 3:00:00 PM", sw.ToString
End Sub
