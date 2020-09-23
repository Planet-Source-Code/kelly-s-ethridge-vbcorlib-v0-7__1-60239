VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PublicFunctions"
Attribute VB_GlobalNameSpace = True
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
'    Module: PublicFunctions
'

''
'
Option Explicit

' Returns accessors to categorized functions.
Public Function Cor() As Constructors
    Set Cor = modPublicFunctions.Cor
End Function

Public Function cArray() As cArray
    Set cArray = modPublicFunctions.cArray
End Function

Public Function cString() As cString
    Set cString = modPublicFunctions.cString
End Function

Public Function comparer() As ComparerStatic
    Set comparer = modPublicFunctions.comparer
End Function

Public Function Environment() As Environment
    Set Environment = modPublicFunctions.Environment
End Function

Public Function BitArray() As BitArrayStatic
    Set BitArray = modPublicFunctions.BitArray
End Function

Public Function Buffer() As Buffer
    Set Buffer = modPublicFunctions.Buffer
End Function

Public Function NumberFormatInfo() As NumberFormatInfoStatic
    Set NumberFormatInfo = modPublicFunctions.NumberFormatInfo
End Function

Public Function BitConverter() As BitConverter
    Set BitConverter = modPublicFunctions.BitConverter
End Function

Public Function Version() As VersionStatic
    Set Version = modPublicFunctions.Version
End Function

Public Function TimeSpan() As TimeSpanStatic
    Set TimeSpan = modPublicFunctions.TimeSpan
End Function

Public Function cDateTime() As cDateTimeStatic
    Set cDateTime = modPublicFunctions.cDateTime
End Function

Public Function DateTimeFormatInfo() As DateTimeFormatInfoStatic
    Set DateTimeFormatInfo = modPublicFunctions.DateTimeFormatInfo
End Function

Public Function CultureInfo() As CultureInfoStatic
    Set CultureInfo = modPublicFunctions.CultureInfo
End Function

Public Function TimeZone() As TimeZoneStatic
    Set TimeZone = modPublicFunctions.TimeZone
End Function

Public Function Path() As Path
    Set Path = modPublicFunctions.Path
End Function

Public Function Encoding() As EncodingStatic
    Set Encoding = modPublicFunctions.Encoding
End Function

Public Function TextReader() As TextReaderStatic
    Set TextReader = modPublicFunctions.TextReader
End Function

Public Function Directory() As Directory
    Set Directory = modPublicFunctions.Directory
End Function

Public Function file() As file
    Set file = modPublicFunctions.file
End Function

Public Function Stream() As StreamStatic
    Set Stream = modPublicFunctions.Stream
End Function

Public Function GregorianCalendar() As GregorianCalendarStatic
    Set GregorianCalendar = modPublicFunctions.GregorianCalendar
End Function

Public Function JulianCalendar() As JulianCalendarStatic
    Set JulianCalendar = modPublicFunctions.JulianCalendar
End Function

Public Function HebrewCalendar() As HebrewCalendarStatic
    Set HebrewCalendar = modPublicFunctions.HebrewCalendar
End Function

Public Function KoreanCalendar() As KoreanCalendarStatic
    Set KoreanCalendar = modPublicFunctions.KoreanCalendar
End Function

Public Function HijriCalendar() As HijriCalendarStatic
    Set HijriCalendar = modPublicFunctions.HijriCalendar
End Function

Public Function Calendar() As CalendarStatic
    Set Calendar = modPublicFunctions.Calendar
End Function

Public Function Console() As Console
    Set Console = modPublicFunctions.Console
End Function

Public Function ArrayList() As ArrayListStatic
    Set ArrayList = modPublicFunctions.ArrayList
End Function

Public Function Registry() As Registry
    Set Registry = modPublicFunctions.Registry
End Function

Public Function RegistryKey() As RegistryKeyStatic
    Set RegistryKey = modPublicFunctions.RegistryKey
End Function

Public Function LongToChar(ByVal value As Long) As Integer
    If value < -32767 Or value > 65535 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value must be between -32767 and 65535 inclusively.", "Value", value)
    If value > 32767 Then value = value - 65536
    LongToChar = value
End Function

Public Function CharToLong(ByVal value As Integer) As Long
    AsWord(CharToLong) = value
End Function

