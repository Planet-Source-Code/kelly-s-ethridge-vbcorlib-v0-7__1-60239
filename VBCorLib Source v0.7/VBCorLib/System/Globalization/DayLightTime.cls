VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DayLightTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: DayLightTime
'

''
' Represents the start and end times for a daylight savings period, and the
' difference between daylight savings time and standard time.
'
Option Explicit
Implements cObject

Private mDelta As TimeSpan
Private mStartTime As cDateTime
Private mEndTime As cDateTime


''
' Returns the date for when the daylight savings period starts.
'
' @return The date of the start of daylight savings.
'
Public Property Get StartTime() As cDateTime
    Set StartTime = mStartTime
End Property

''
' Returns the date for when the daylight savings period ends.
'
' @return The date of the end of daylight savings.
'
Public Property Get EndTime() As cDateTime
    Set EndTime = mEndTime
End Property

''
' Returns the difference between stadard time and daylight savings.
'
' @return The difference between stadard time and daylight savings.
' @remarks This is the amount to change the current time when switching
' between standard and daylight savings time.
'
Public Property Get Delta() As TimeSpan
    Set Delta = mDelta
End Property

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = App.Title & "." & TypeName(Me)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef value As Variant) As Boolean
    If IsObject(value) Then
        Equals = (value Is Me)
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef StartTime As Variant, ByRef EndTime As Variant, ByVal Delta As TimeSpan)
    Set mDelta = Delta
    Set mStartTime = cDateTime.GetcDateTime(StartTime)
    Set mEndTime = cDateTime.GetcDateTime(EndTime)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set mStartTime = .ReadProperty("StartTime")
        Set mEndTime = .ReadProperty("EndTime")
        Set mDelta = .ReadProperty("Delta")
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "StartTime", mStartTime
        .WriteProperty "EndTime", mEndTime
        .WriteProperty "Delta", mDelta
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(value As Variant) As Boolean
    cObject_Equals = Equals(value)
End Function

Private Function cObject_GetHashcode() As Long
    cObject_GetHashcode = GetHashCode
End Function

Private Function cObject_ToString() As String
    cObject_ToString = ToString
End Function
