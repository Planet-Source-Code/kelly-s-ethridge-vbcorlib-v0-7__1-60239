VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Version"
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
'    Module: Version
'

''
' Represents a version number.
'
Option Explicit
Implements cObject
Implements ICloneable
Implements IComparable

Private Const PROP_MAJOR            As String = "Major"
Private Const PROP_MINOR            As String = "Minor"
Private Const PROP_BUILD            As String = "Build"
Private Const PROP_REVISION         As String = "Revision"

Private Const UNUSED                As Long = -1
Private Const DEF_MAJOR             As Long = 0
Private Const DEF_MINOR             As Long = 0
Private Const DEF_BUILD             As Long = UNUSED
Private Const DEF_REVISION          As Long = UNUSED


Private mMajor As Long
Private mMinor As Long
Private mBuild As Long
Private mRevision As Long


''
' Returns the major version number of this instance.
'
' @return The major version number. This is a non-negative number.
'
Public Property Get Major() As Long
    Major = mMajor
End Property

''
' Returns the minor version number of this instance.
'
' @return The minor version number. This is a non-negative number.
'
Public Property Get Minor() As Long
    Minor = mMinor
End Property

''
' Returns the build version number of this instance.
'
' @return The build version number.
' @remarks The build version is an optional tracking number. If the number
' was not used when this instance was created, then -1 is returned.
'
Public Property Get Build() As Long
    Build = mBuild
End Property

''
' Returns the revision version number of this instance.
'
' @return The revision version number.
' @remarks The build version is an optional tracking number. If the number
' was not used when this instance was created, then -1 is returned.
'
Public Property Get Revision() As Long
    Revision = mRevision
End Property

''
' Returns a cloned object of this instance.
'
' @return The cloned version of this instanced.
'
Public Function Clone() As Object
    Dim ret As New Version
    ret.CloneHelper mMajor, mMinor, mBuild, mRevision
    Set Clone = ret
End Function

''
' Compares this instance to a passed in value.
'
' @param value The value to compare this instance against.
' @return Value indicating the relation of this instance to the passed in value.
' @remarks A negative number indicates this instance is less than the value.
' A positive number indicates this instance is greater than the value.
' Zero indicates this instance is equal to the value.
'
Public Function CompareTo(ByRef value As Variant) As Long
    Dim v As Version
   
    Select Case VarType(value)
        Case vbObject
            If value Is Nothing Then: CompareTo = 1: Exit Function
            If Not TypeOf value Is Version Then Throw Cor.NewArgumentException("Value must of type Version.", "value")
            
            Set v = value
            If Not EqualFields(mMajor, v.Major, CompareTo) Then Exit Function
            If Not EqualFields(mMinor, v.Minor, CompareTo) Then Exit Function
            If Not EqualFields(mBuild, v.Build, CompareTo) Then Exit Function
            If Not EqualFields(mRevision, v.Revision, CompareTo) Then Exit Function
        
        Case vbNull
            CompareTo = 1
        Case Else
            Throw Cor.NewArgumentException("Value must of type Version.", "value")
    End Select
End Function

''
' Returns a string representation of this object instance.
'
' @param vFieldCount The number of fields to be included in the result string.
' @return String representing this instance in the form of Major.Minor[.Build[.Revision]]
'
Public Function ToString(Optional ByRef vFieldCount As Variant) As String
    Dim ret As String
    Dim fieldcount As Long
    Dim maxfield As Long
    
    If IsMissing(vFieldCount) Then
        fieldcount = -1
    Else
        fieldcount = vFieldCount
        maxfield = GetMaxFieldCount
        If fieldcount < 0 Or fieldcount > maxfield Then Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_VersionFieldCount, maxfield), "vFieldCount", fieldcount)
    End If
    
    If fieldcount <> 0 Then ret = mMajor
    If fieldcount < 0 Or fieldcount > 1 Then ret = ret & "." & mMinor
    If mBuild >= 0 And (fieldcount < 0 Or fieldcount > 2) Then
        ret = ret & "." & mBuild
        If mRevision >= 0 And (fieldcount < 0 Or fieldcount > 3) Then ret = ret & "." & mRevision
    End If
    ToString = ret
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef value As Variant) As Boolean
    If IsObject(value) Then
        If TypeOf value Is Version Then
            Equals = (CompareTo(value) = 0)
        End If
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    Dim ret As Long
    ret = Helper.ShiftLeft(mMajor And &HF, 28)
    ret = ret Or Helper.ShiftLeft(mMinor And &HFF, 20)
    ret = ret Or Helper.ShiftLeft(mBuild And &HFF, 12)
    GetHashCode = ret Or (mRevision And &HFFF&)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Major As Long, ByVal Minor As Long, ByRef vBuild As Variant, ByRef vRevision As Variant)
    If Major < 0 Or Minor < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Version's parameters must be equal to or greater than zero.")
    
    If Not IsMissing(vBuild) Then
        mBuild = vBuild
        If mBuild < 0 Then Throw Cor.NewArgumentOutOfRangeException("Version's parameters must be equal to or greater than zero.")
    End If
    If Not IsMissing(vRevision) Then
        If mBuild < 0 Then Throw Cor.NewArgumentException("Cannot include a revision without a build value.", "Revision")
        mRevision = vRevision
        If mRevision < 0 Then Throw Cor.NewArgumentOutOfRangeException("Version's parameters must be equal to or greater than zero.")
    End If
    
    mMajor = Major
    mMinor = Minor
End Sub

Friend Sub CloneHelper(ByVal Major As Long, ByVal Minor As Long, ByVal Build As Long, ByVal Revision As Long)
    mMajor = Major
    mMinor = Minor
    mBuild = Build
    mRevision = Revision
End Sub

Friend Sub Parse(ByRef s As String)
    Dim parts() As String
    Dim ub As Long
    
    parts = Split(s, ".")
    ub = UBound(parts)
    If ub < 1 Or ub > 3 Then _
        Throw Cor.NewArgumentException("String must contain 2 to 4 fields.")
    mMajor = parts(0)
    mMinor = parts(1)
    If mMajor < 0 Or mMinor < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Version's parameters must be equal to or greater than zero.")
    
    If ub > 1 Then
        mBuild = parts(2)
        If mBuild < 0 Then Throw Cor.NewArgumentOutOfRangeException("Version's parameters must be equal to or greater than zero.")
        If ub > 2 Then
            mRevision = parts(3)
            If mRevision < 0 Then Throw Cor.NewArgumentOutOfRangeException("Version's parameters must be equal to or greater than zero.")
        End If
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetMaxFieldCount() As Long
    If mBuild < 0 Then
        GetMaxFieldCount = 2
    ElseIf mRevision < 0 Then
        GetMaxFieldCount = 3
    Else
        GetMaxFieldCount = 4
    End If
End Function

Private Function EqualFields(ByVal MyField As Long, ByVal TheirField As Long, ByRef retval As Long) As Boolean
    If MyField < TheirField Then
        retval = -1
    ElseIf MyField > TheirField Then
        retval = 1
    Else
        retval = 0
        EqualFields = True
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    mBuild = UNUSED
    mRevision = UNUSED
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mMajor = .ReadProperty(PROP_MAJOR, DEF_MAJOR)
        mMinor = .ReadProperty(PROP_MINOR, DEF_MINOR)
        mBuild = .ReadProperty(PROP_BUILD, DEF_BUILD)
        mRevision = .ReadProperty(PROP_REVISION, DEF_REVISION)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PROP_MAJOR, mMajor
        .WriteProperty PROP_MINOR, mMinor
        .WriteProperty PROP_BUILD, mBuild
        .WriteProperty PROP_REVISION, mRevision
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparable_CompareTo(value As Variant) As Long
    IComparable_CompareTo = CompareTo(value)
End Function
