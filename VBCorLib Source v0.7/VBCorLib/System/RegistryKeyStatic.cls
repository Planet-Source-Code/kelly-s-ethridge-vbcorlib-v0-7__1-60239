VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKeyStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: RegistryKeyStatic
'

Option Explicit

Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, ByRef phkResult As Long) As Long

Public Enum RegistryHive
    ClassesRoot = &H80000000
    CurrentConfig = &H80000005
    CurrentUser = &H80000001
    DynData = &H80000006
    LocalMachine = &H80000002
    Users = &H80000003
    PerformanceData = &H80000004
End Enum



Public Function OpenRemoteBaseKey(ByVal hKey As RegistryHive, ByVal MachineName As String) As RegistryKey
    Dim e As Long
    Dim NewKey As Long
    
    e = RegConnectRegistry(MachineName, hKey, NewKey)
    If e <> NO_ERROR Then IOError e, "OpenRemoteBaseKey"
    
    Set OpenRemoteBaseKey = New RegistryKey
    OpenRemoteBaseKey.InitRemote NewKey
End Function
