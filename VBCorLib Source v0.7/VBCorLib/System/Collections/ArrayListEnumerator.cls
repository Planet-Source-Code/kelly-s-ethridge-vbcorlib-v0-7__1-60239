VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayListEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: ArrayListEnumerator
'
Option Explicit
Implements cObject
Implements IEnumerator
Implements ICloneable

Private mList As ArrayList
Private mStartIndex As Long
Private mEndIndex As Long
Private mIndex As Long
Private mVersion As Long



' This is to allow code such as the following to function properly.
'
' For Each v In obj.GetEnumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function


' Called when being initialized by an ArrayList object.
Friend Sub Init(ByVal list As ArrayList, ByVal Index As Long, ByVal Count As Long)
    Set mList = list
    mVersion = list.Version
    mStartIndex = Index
    mEndIndex = Index + Count - 1
    mIndex = Index - 1
End Sub

Friend Sub CloneHelper(ByVal list As ArrayList, ByVal startindex As Long, ByVal Index As Long)
    Init list, startindex, list.Count
    mIndex = Index
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Dim ret As New ArrayListEnumerator
    ret.CloneHelper mList, mStartIndex, mIndex
    Set ICloneable_Clone = ret
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex < mStartIndex Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mIndex > mEndIndex Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
        
    Helper.MoveVariant IEnumerator_Current, mList(mIndex)
End Property

Private Function IEnumerator_MoveNext() As Boolean
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex > mEndIndex Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
    If mIndex = mEndIndex Then
        mIndex = mEndIndex + 1
    Else
        mIndex = mIndex + 1
        IEnumerator_MoveNext = True
    End If
End Function

Private Sub IEnumerator_Reset()
    mIndex = mStartIndex - 1
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(value As Variant) As Boolean
    If IsObject(value) Then
        cObject_Equals = (value Is Me)
    End If
End Function

Private Function cObject_GetHashcode() As Long
    cObject_GetHashcode = ObjPtr(Me)
End Function

Private Function cObject_ToString() As String
    cObject_ToString = App.Title & "." & TypeName(Me)
End Function

