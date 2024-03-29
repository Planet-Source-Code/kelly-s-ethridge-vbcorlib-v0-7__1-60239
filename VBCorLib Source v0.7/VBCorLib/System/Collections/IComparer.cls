VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: IComparer
'

''
' A standard interface used to implement custom comparers.
'
' @remarks<br>
' Use this interface to create a comparer that can be used in searching and sorting
' functions in the collections. Generally, custom comparers are used for special-case
' values to compare, such as user-defined types. In order to not have to write logic
' to compare all types of values, the method could delegate general cases to the
' default comparer, and compare only the specialized values.
'
Option Explicit

''
' Compares two elements and returns their relationship.
'
' @param x The left-hand-side of the equation.
' @param y the right-hand-side of the equation.
' @return Value indicating the relationships between the values.
' @remarks <br>
' Comparison guide:<br>
' x is the same value as y returns a 0.<br>
' x is less than the value y returns a negative value (-1).<br>
' x is greater than the value y returns a positive value (1).<br>
'
Public Function Compare(ByRef x As Variant, ByRef y As Variant) As Long: End Function

