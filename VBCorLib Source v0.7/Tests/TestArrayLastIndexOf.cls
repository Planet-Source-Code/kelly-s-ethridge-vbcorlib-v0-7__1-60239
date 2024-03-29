VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayLastIndexOf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestLong()
    Dim a() As Long
    Dim i As Long
    
    ReDim a(-1 To 100)
    For i = 0 To UBound(a)
        a(i) = i
    Next i
    
    i = carray.LastIndexOf(a, CLng(UBound(a)))
    
    Assert.Equals UBound(a), i, "LastIndexOf should return " & UBound(a)
End Sub

Public Sub TestString()
    Dim a() As String
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To UBound(a)
        a(i) = "Item " & i
    Next i
    
    i = carray.LastIndexOf(a, "Item " & UBound(a))
    Assert.Equals UBound(a), i, "LastIndexOf should return " & UBound(a)
End Sub

Public Sub TestDate()
    Dim a() As Date
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = DateSerial(i, 1, 1)
    Next i
    
    i = carray.LastIndexOf(a, DateSerial(59, 1, 1))
    Assert.Equals 59, i, "LastIndexOf should return 59"
End Sub

Public Sub TestVariant()
    Dim a() As Variant
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = "Item " & i
    Next i
    Set a(88) = Me
    
    i = carray.LastIndexOf(a, Me)
    Assert.Equals 88, i, "LastIndexOf should return 88"
End Sub

Public Sub TestDefaultComparer()
    Dim a() As String
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = "Item " & i
    Next i
    
    i = carray.LastIndexOf(a, "Item 88", , , comparer.Default)
    Assert.Equals 88, i, "LastIndexOf should return 88."
End Sub

Public Sub TestNotInSubArray()
    Dim a() As Long
    Dim i As Long
    
    ReDim a(-1 To 100)
    For i = 0 To 100
        a(i) = i
    Next i
    
    i = carray.LastIndexOf(a, 50&, 20, 20)
    Assert.Equals -2, i, "LastIndexOfEx should return -1."
End Sub

Public Sub TestNothing()
    Dim a() As Variant
    Dim i As Long
    
    ReDim a(-10 To 100)
    For i = -10 To 100
        a(i) = "Item " & i
    Next i
    Set a(LBound(a)) = Nothing
    i = carray.LastIndexOf(a, Nothing)
    Assert.Equals LBound(a), i, "LastIndexOf should return " & LBound(a)
End Sub

