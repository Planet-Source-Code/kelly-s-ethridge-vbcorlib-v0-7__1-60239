VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryReader"
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
'    Module: BinaryReader
'

''
' Provides a set of methods to read typed information from a Stream.
'
' @remarks
' <p>The bytes from a stream can be read and cast to a specific datatype. Each
' datatype will determine the number of bytes to be read from the stream. Once
' the number of bytes have been read, they are assembled into the requested
' datatype, advancing the position in the stream the required number of bytes.</p>
' <p>Characters can be read from the Stream using the provided Encoding system.
' The default encoding system is equal to New UTF8Encoding.</p>
'
Option Explicit
Implements cObject

Private Const DEF_CAPACITY  As Long = 16
Private Const BYTE_ARRAY    As Long = vbByte Or vbArray
Private Const CHAR_ARRAY    As Long = vbInteger Or vbArray
Private Const MIN_DATE      As Double = -657434#
Private Const MAX_DATE      As Double = 2958465.99998843


Private mStream As Stream
Private mEncoding As Encoding
Private mIsOpen As Boolean
Private mBuffer() As Byte
Private mCapacity As Long
Private mPtrBuffer As Long
Private mOneChar(0) As Integer


''
' Returns the underlying Stream being read by the Reader.
'
' @return The underlying Stream.
'
Public Property Get BaseStream() As Stream
    Set BaseStream = mStream
End Property

''
' Closes the reader and underlying Stream.
'
Public Sub CloseReader()
    mStream.CloseStream
End Sub

''
' Reads a byte from the stream.
'
' @return The byte read from the stream, or -1 if there are no more bytes in the stream.
'
Public Function Read() As Long
    VerifyIsOpen
    If ReadCharBytes(mOneChar, 0, 1) = 0 Then
        Read = -1
    Else
        AsWord(Read) = mOneChar(0)
    End If
End Function

''
' Fills either a Byte array or Integer array with the specified number of elements.
'
' @param Buffer The array to write the data to.
' @param Index The first index in <i>Buffer</i> to start writing data.
' @param Count The number of elements to be written to the array.
' @return The actual number of elements written to the array. This could be less than the requested amount.
' @remarks If a Byte array is passed in, then the bytes from the underlying stream are
' copied to the array. If an Integer array is passed in, then enough bytes are read from
' the stream to produce the requested number of decoded characters. The decoded characters
' are placed in the array starting at <i>Index</i>.
'
Public Function ReadBuffer(ByRef Buffer As Variant, ByVal Index As Long, ByVal Count As Long) As Long
    Dim Bytes() As Byte
    Dim Chars() As Integer
    
    On Error GoTo errTrap
    Select Case VarType(Buffer)
        Case BYTE_ARRAY
            SAPtr(Bytes) = cArray.GetValidArrayPointer(Buffer)
            ReadBuffer = mStream.ReadBlock(Bytes, Index, Count)
        Case CHAR_ARRAY
            SAPtr(Chars) = cArray.GetValidArrayPointer(Buffer)
            ReadBuffer = ReadCharBytes(Chars, Index, Count)
        Case Else
            Throw Cor.NewArgumentException("A Byte or Integer array is required.", "Buffer")
    End Select
    
errTrap:
    SAPtr(Bytes) = 0
    SAPtr(Chars) = 0
End Function

''
' Returns the next character to be decoded from the byte stream. The position
' in the stream is not advanced after the read.
'
' @return The next character to be Read, or -1 if no more characters are in the stream.
' @remarks This function requires the stream to support Seeking. If it does not,
' then -1 will always be returned.
'
Public Function PeekChar() As Long
    Dim pos As Long
    
    VerifyIsOpen
    
    If mStream.CanSeek Then
        pos = mStream.Position
        PeekChar = Read
        mStream.Position = pos
    Else
        PeekChar = -1
    End If
End Function

''
' Reads the next byte in the stream, and advances the position one byte.
'
' @return The next byte in the stream.
Public Function ReadByte() As Byte
    FillBuffer 1
    ReadByte = mBuffer(0)
End Function

''
' Reads the next decode character in the stream, and advances the position
' the number of bytes requirece to assemble a single character.
'
' @return The next character in the stream.
' @remarks If no characters are left in the stream, and EndOfStreamException is thrown.
'
Public Function ReadChar() As Long
    ReadChar = Read
    If ReadChar = -1 Then Throw New EndOfStreamException
End Function

''
' Reads the specified number of bytes from the stream and returns them in an array.
'
' @param Count The number of bytes to read from the stream.
' @return A byte array containing the bytes read from the stream.
' @remarks If not enough bytes are in the stream, then the remaining
' bytes are returned. If there are no bytes in the stream, then an
' empty array is returned, not a null array.
'
Public Function ReadBytes(ByVal Count As Long) As Byte()
    Dim ret() As Byte
    Dim num As Long
    
    VerifyIsOpen
    
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    
    If Count > 0 Then
        ReDim ret(0 To Count - 1)
        num = mStream.ReadBlock(ret, 0, Count)
        If num = 0 Then Exit Function
        If num <> Count Then ReDim Preserve ret(0 To num - 1)
    Else
        ret = cArray.CreateInstance(ciByte)
    End If
    ReadBytes = ret
End Function

''
' Returns a specified number of characters decoded from the stream.
'
' @param Count The number of characters to decode from the stream.
' @return The array of characters decoded from the stream.
' @remarks If not enough characters could be decoded from the stream,
' then only the remaining characters are returned. If no characters
' could be decoded, then an empty array is returned, not a null array.
'
Public Function ReadChars(ByVal Count As Long) As Integer()
    Dim ret() As Integer
    Dim CharsRead As Long
    
    VerifyIsOpen
    
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    
    If Count > 0 Then
        ReDim ret(0 To Count - 1)
        CharsRead = ReadCharBytes(ret, 0, Count)
        If CharsRead = 0 Then Throw New EndOfStreamException
        If Count <> CharsRead Then ReDim Preserve ret(0 To CharsRead - 1)
    Else
        ret = cArray.CreateInstance(ciInteger)
    End If
    ReadChars = ret
End Function

''
' Returns a boolean from the stream.
'
' @return A boolean value.
' @remarks A single byte is read from the stream. If the byte
' is non-zero, then True is returned, otherwise False.
'
Public Function ReadBoolean() As Boolean
    VerifyIsOpen
    ReadBoolean = (ReadByte <> 0)
End Function

''
' Reads 4 bytes from the stream and returns them as a Long datatype.
'
' @return A Long datatype read from the stream.
'
Public Function ReadLong() As Long
    FillBuffer 4
    ReadLong = MemLong(mPtrBuffer)
End Function

''
' Reads 2 bytes from the stream and returns them as an Integer datatype.
'
' @return An Integer datatype read from the stream.
'
Public Function ReadInteger() As Integer
    FillBuffer 2
    ReadInteger = MemWord(mPtrBuffer)
End Function

''
' Reads 8 bytes from the stream and returns them as a Double datatype.
'
' @return A Double datatype read from the stream.
'
Public Function ReadDouble() As Double
    FillBuffer 8
    ReadDouble = AsDouble(ByVal mPtrBuffer)
End Function

''
' Reads 8 bytes from the stream and returns them as a Date datatype.
'
' @return A Date datatype read from the stream.
' @remarks It is possible to read a combination of 8 bytes that
' convert into an invalid Date. This will cause and Overflow error.
'
Public Function ReadDate() As Date
    Dim d As Double
    FillBuffer 8
    d = AsDouble(ByVal mPtrBuffer)
    If d < MIN_DATE Or d > MAX_DATE Then Throw New OverflowException
    ReadDate = d
End Function

''
' Reads 8 bytes from the stream and returns them as a Currency datatype.
'
' @return A Currency datatype.
'
Public Function ReadCurrency() As Currency
    FillBuffer 8
    ReadCurrency = AsCurr(ByVal mPtrBuffer)
End Function

''
' Reads 4 bytes from the stream and returns them as a Single datatype.
'
' @return A Single datatype.
'
Public Function ReadSingle() As Single
    FillBuffer 4
    ReadSingle = AsSingle(ByVal mPtrBuffer)
End Function

''
' Reads 16 bytes from the stream and returns them as a Variant Decimal datatype.
'
' @return A Variant Decimal datatype.
' @remarks <p>The Decimal datatype is created from 16 bytes with the same
' layout as used in .NET. The layout is different than that of VB.<br><br>
' VB Decimal layout<br>
' bytes 0-1: Variant datatype information<br>
' bytes 2: precision<br>
' bytes 3: sign (&h80 is negative)<br>
' bytes 4-7: the 32 highest bits in the 96bit value<br>
' bytes 8-11: the 32 lowest bits in the 96bit value<br>
' bytes 12-15: the 32 middle bits in the 96bit value<br>
' <br>
' .NET Decimal layout<br>
' bytes 0-3: the 32 lowest bits in the 96bit value<br>
' bytes 4-7: the 32 middle bits in the 96bit value<br>
' bytes 8-11: the 32 highest bits in the 96bit value<br>
' bytes 12-13: unused (zero)<br>
' bytes 14: precision<br>
' bytes 15: sign (&h80 is negative)<br><br>
' If the sign byte is non-zero, then the Decimal will be set as negative.<br>
'
Public Function ReadDecimal() As Variant
    FillBuffer 16
    ReadDecimal = BitConverter.ToDecimal(mBuffer, 0)
End Function

''
' Reads a String from the stream.
'
' @return The string datatype.
' @remarks The string is stored in the stream with the number of encoded bytes preceding
' the actual string data. The bytes are not the typical 4 bytes as is used by VB when
' representing the length of the string. The bytes represent a value encoded as
' 7bits per byte. While the next read byte has the high bit set (&h80), then the
' following byte is also part of the length value. For each byte read, first 7 bits of
' that byte is pushed out 7 bits multiplied by the current byte number - 1 in the sequence.
' This is not the normal shifting of the current sum of the values. Each read byte
' must be shifted left individually, as each byte represents a higher set of bits
' in the resulting number.<br><br>
' We AND the byte with &h3f because we only want the lower 7 bits.<br>
' byte 1: (byte and &h3f) << shifted 0 bits added to the sum<br>
' byte 2: (byte and &h3f) << shifted 7 bits added to the sum<br>
' byte 3: (byte and &H3f) << shifted 14 bits added to the sum<br><br>
' .. this continues until a byte is less than 128. At which point, it is shifted and summed like the
' rest, but no more bytes are to be read in. The sum now contains the number of bytes to
' be read in and decoded into the string. The same Encoding type must be used to retrieve
' the correct string value. Using a different Encoding method will create the wrong string.
'
' <p>Once the length is determined, then enough bytes are read in that can be decoded
' to the string using the current Encoding system.</p>
'
Public Function ReadString() As String
    Dim l As Long
    Dim b() As Byte
    Dim n As Long
    
    VerifyIsOpen
    l = ReadStringLength
    If l > 0 Then
        ReDim b(0 To l - 1)
        n = mStream.ReadBlock(b, 0, l)
        If n < l Then Throw New EndOfStreamException
        ReadString = mEncoding.GetString(b)
    End If
End Function

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
' @param value The value to compare equalit to.
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
Friend Sub Init(ByVal Stream As Stream, ByVal Encoding As Encoding)
    If Stream Is Nothing Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Stream), "Stream")
    If Not Stream.CanRead Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ReadableStreamRequired), "Stream")
    
    Set mStream = Stream
    If Encoding Is Nothing Then
        Set mEncoding = New UTF8Encoding
    Else
        Set mEncoding = Encoding
    End If
    mIsOpen = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyIsOpen()
    If Not mIsOpen Then Throw Cor.NewObjectDisposedException("BinaryReader", "Cannot read from a closed BinaryReader.")
End Sub

Private Sub FillBuffer(ByVal Count As Long)
    Dim num As Long
    
    VerifyIsOpen
    
    ' optimize for reading a single byte
    If Count = 1 Then
        num = mStream.ReadByte
        If num >= 0 Then
            mBuffer(0) = num
            Exit Sub
        End If
        num = 0
    Else
        num = mStream.ReadBlock(mBuffer, 0, Count)
    End If
    If num = 0 Then Throw New EndOfStreamException
End Sub

Private Function ReadStringLength() As Long
    Dim ret As Long
    Dim num As Long
    Dim shiftCount As Long
    
    num = mStream.ReadByte
    Do While num >= &H80
        ret = ret + Helper.ShiftLeft(num And &H7F, shiftCount)
        num = mStream.ReadByte
        shiftCount = shiftCount + 7
    Loop
    If num > 0 Then
        ReadStringLength = ret + Helper.ShiftLeft(num, shiftCount)
    Else
        ReadStringLength = ret
    End If
End Function

Private Function ReadCharBytes(ByRef Chars() As Integer, ByVal Index As Long, ByVal Count As Long) As Long
    Dim numread As Long
    Dim pos As Long
    Dim b As Long
    
    If Index < LBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "Index", Index)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If Index + Count - 1 > UBound(Chars) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    EnsureCapacity mEncoding.GetMaxByteCount(Count)
    
    Do While numread < Count
        b = mStream.ReadByte
        If b < 0 Then Exit Do
        mBuffer(pos) = b
        pos = pos + 1
        numread = numread + mEncoding.GetCharsEx(mBuffer, 0, pos, Chars, Index)
    Loop
    ReadCharBytes = pos
End Function

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    Dim NewCapacity As Long
    
    If RequiredCapacity <= mCapacity Then Exit Sub
    NewCapacity = mCapacity * 2
    If NewCapacity < RequiredCapacity Then NewCapacity = RequiredCapacity
    ReDim Preserve mBuffer(0 To NewCapacity - 1)
    mCapacity = NewCapacity
    mPtrBuffer = VarPtr(mBuffer(0))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    EnsureCapacity DEF_CAPACITY
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
