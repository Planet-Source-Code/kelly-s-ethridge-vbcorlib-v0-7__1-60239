VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryStream"
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
'    Module: MemoryStream
'

''
' Provides an in-memory stream of data that can be read/written and modified
' using Stream methods.
'
' @remarks <p>The MemoryStream uses an array of bytes to back the data
' in memory. When the object is first created, no array is allocated until the
' first time the data must be accessed. The default capacity of the array is 4096 bytes.
' The capacity will grow by doubling the size of the current capacity.
' If The capacity needs to be larger, consider setting the Capacity directly to prevent
' the array from having to be resized as new data is written to the stream.</p>
' <p>If the MemoryStream is created using a supplied Byte array, then no internal array
' will be allocated. The MemoryStream will access the supplied Byte array as if it
' were created by the MemoryStream object. As the underlying array is modified, the
' changes can be seen by directly accessing the array. Also, by modifying the array
' directly, the underlying stream will reflect the new data.</p>
' <p><b>CAUTION:</b> If a Byte array is being supplied to the MemoryStream, then the
' array must outlive the MemoryStream object. It is the responsibility of the user to
' destroy the array after the stream has been destroyed. The MemoryStream object
' maintains a reference to the same array in memory, but does not deallocate the
' array when finished. While the object is alive, the array is locked to prevent
' it from being ReDimed or Erased. Closing the stream will release the lock and
' reference to the array, at which point the user has full control of the array.</p>
' <p>The Currency datatypes used are there for consistency across all stream objects.
' Internally, the MemoryStream only supports 2 ^ 31 bytes. Other streams may support more.</p>
'
' <br><h4>Constructors:</h4><p>
' <font face="Courier New">
' Set obj = New MemoryStream<br>
' Set obj = NewMemoryStream(ByteArray, [Index], [Count])
' Set obj = Cor.NewMemoryStream(ByteArray, [Index], [Count])
' </font><br></p>
Option Explicit
Implements cObject
Implements Stream

Private Const DEF_CAPACITY As Long = 4096

Private mBufferSet As Boolean
Private mBuffer() As Byte
Private mCanWrite As Boolean
Private mPosition As Long
Private mCapacity As Long
Private mLength As Long
Private mIsClosed As Boolean
Private mIsUserArray As Boolean
Private mIndex As Long



''
' Begins an asynchronous buffer read. For a MemoryStream, calling an asynchronous buffer
' read is identical to calling ReadBlock.
'
' @param Buffer The destination array to copy the bytes into.
' @param Offset The index in Buffer to begin copying to.
' @param Count The maximum number of bytes to copy.
' @param Callback A listening object to be notified that the read is finished.
' @param State User data that can be carried within the IAsyncResult object returned.
' @return An IAsyncResult object used to identify the corrisponding Endread and report
' information about the buffer read.
' @remarks <p>Since the stream is directly stored and accessed in memory, there is no
' waiting for reads to occur. The BeginRead method works identical to the Read method
' in terms of immediately reading stream data into the <i>Buffer</i>.</p>
' <p>When defining the Count for the number of maximum bytes to be read, if the
' stream does not have enough bytes, the read will copy less bytes than Count.</p>
' <p>If an exception has occured during the read, it will not be thrown until the
' EndRead is executed.
'
Public Function BeginRead(ByRef Buffer() As Byte, ByVal offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Dim ret As StreamAsyncResult
    
    Set ret = Cor.NewStreamAsyncResult(State)
    On Error GoTo errTrap
    With ret
        .BytesRead = ReadBlock(Buffer, offset, Count)
        .IsCompleted = True
        .IsReadType = True
    End With
    
    If Not Callback Is Nothing Then Callback.Execute ret

errTrap:
    Dim ex As Exception
    If Catch(ex, Err) Then Set ret.Exception = ex
    Set BeginRead = ret
End Function

''
' Begins an asynchronous buffer write. For a MemoryStream, having an asynchronous buffer
' write is of no use.
'
' @param Buffer The source array to copy the bytes from into the stream.
' @param Offset The starting index in the source <i>Buffer</i> to begin copying from.
' @param Count The maximum number of bytes to be copied into the stream.
' @param Callback A user supplied object to be notified of the finish of the writing.
' @param State User data that can be carried within the IAsyncResult object return.
' @return An IAsyncResult object used to identify the corrisponding EndBread and report
' information about the buffer read.
' @remarks <p>Since the stream is directly stored and accessed in memory, there is no
' waiting for reads to occur. The BeginWrite method works identical to the WriteBlock method
' in terms of reading immediately writing stream data to the stream.</p>
' <p>When defining the Count for the number of maximum bytes to be written, if the
' <i>Buffer</i> does not have enough bytes, an exception will be thrown in the EndWrite method.</p>
' <p>If an exception has occured during the read, it will not be thrown until the
' EndRead is executed.
'
Public Function BeginWrite(ByRef Buffer() As Byte, ByVal offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Dim ret As StreamAsyncResult
    
    Set ret = Cor.NewStreamAsyncResult(State)
    On Error GoTo errTrap:
    With ret
        .CompletedSynchronously = True
        .IsCompleted = True
    End With
    WriteBlock Buffer, offset, Count
    
    If Not Callback Is Nothing Then Callback.Execute ret
    
errTrap:
    Dim ex As Exception
    If Catch(ex, Err) Then Set ret.Exception = ex
    Set BeginWrite = ret
End Function

''
' Signifies the end of an asynchronous read from the stream.
'
' @param AsyncResult The IAsyncResult object returned from the BeginRead function.
' @return The number of bytes read from the stream in to <i>Buffer</i> in the BeginRead method.
' @remarks <p>This method is generally called anytime after BeginRead for a MemoryStream.
' Since the MemoryStream uses memory for the data store, there is no need to wait for
' the data to arrive. The BeginRead/EndRead calls are redundant and should not be
' used with the MemoryStream. Instead use the Read Method.</p>
' <p>If any exceptions were caused during the reading then they will be thrown
' during the call to EndRead.</p>
'
Public Function EndRead(ByVal AsyncResult As IAsyncResult) As Long
    Dim result As StreamAsyncResult
    
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult object is required.", "AsyncResult")
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    Set result = AsyncResult
    If Not result.IsReadType Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    If result.EndCalled Then _
        Throw Cor.NewInvalidOperationException("The EndRead has already been called.")
    
    If Not result.Exception Is Nothing Then Throw result.Exception
    
    EndRead = result.BytesRead
End Function

''
' Signifies the end of an asynchronous write to the stream.
'
' @param AsyncResult The IAsyncResult returned from the BeginWrite function.
' @remarks <p>This method is generally called anytime after BeginWrite for a MemoryStream.
' Since the MemoryStream uses memory for the data store, there is no need to wait for
' the data to be written. The BeginWrite/EndWrite calls are redundant and should not be
' used with the MemoryStream. Instead use the WriteBlock Method.</p>
' <p>If any exceptions were caused during the reading then they will be thrown
' during the call to EndRead.</p>
'
Public Sub EndWrite(ByVal AsyncResult As IAsyncResult)
    Dim result As StreamAsyncResult
    
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult object is required.", "AsyncResult")
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    Set result = AsyncResult
    If result.IsReadType Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    If result.EndCalled Then _
        Throw Cor.NewInvalidOperationException("The EndRead has already been called.")

    If Not result.Exception Is Nothing Then Throw result.Exception
End Sub

''
' Determines if the current MemoryStream can be read.
'
' @return The status of the ability to read the stream.
' @remarks If the MemoryStream is open, then it can be read. If it is closed, it cannot be read.
'
Public Property Get CanRead() As Boolean
    CanRead = Not mIsClosed
End Property

''
' Determines if the current MemoryStream can be written to.
'
' @return The status of the ability to write to the stream.
' @remarks If the stream was created from a user-supplied array, then the user
' will determine if the stream is allowed to be written to. Otherwise, the stream
' is always writable while it is open.
'
Public Property Get CanWrite() As Boolean
    CanWrite = mCanWrite And (Not mIsClosed)
End Property

''
' Determines if the current MemoryStream can seek positions within the stream.
'
' @return The status of the ability to seek a position within the stream.
' @remarks If the MemoryStream is open, then it is seekable, otherwise it is not.
'
Public Property Get CanSeek() As Boolean
    CanSeek = Not mIsClosed
End Property

''
' Returns the current capacity of the internal stream buffer.
'
' @return The number of bytes the internal array contains.
' @remarks If the the capacity is reached while writing, then the capacity
' will be expanded, unless the underlying byte array is user-supplied. In
' this case, an exception will be thrown if the capacity is changed.
'
Public Property Get capacity() As Long
    capacity = mCapacity
End Property

''
' Sets the current capacity for the internal stream buffer.
'
' @param RHS The number of bytes to allocate for the internal array.
' @remarks The capacity cannot be set less than the current length of
' the stream. If the underlying byte array is user-supplied, then an
' exception will be thrown if the capacity is modified.
'
Public Property Let capacity(ByVal RHS As Long)
    VerifyOpen
    
    If mIsUserArray Then _
        Throw Cor.NewNotSupportedException("The capacity cannot be modified when using a user supplied byte array.")
    If RHS < mLength Then _
        Throw Cor.NewArgumentOutOfRangeException("Capacity cannot be set less than then length of the stream.", "Capacity", RHS)
    
    If RHS = mLength Then Exit Property
    mCapacity = RHS - 1
    If mBufferSet Then
        ReDim Preserve mBuffer(0 To mCapacity - 1)
    Else
        VerifyBuffer
    End If
End Property

''
' Returns the current length of the stream.
'
' @return The number of bytes the current stream is.
' @remarks The length of the stream differs from the capacity in that the
' capacity is the amount of allocated memory used to store the stream, and the
' length is the number of bytes used from within the allocated memory.
'
Public Property Get Length() As Currency
    Length = mLength
End Property

''
' Returns the current position within the stream.
'
' @return The byte index within the stream.
' @remarks The position is the index at which the next read or write
' will occur. Position is alway zero based, even if a user-supplied
' array is not.
'
Public Property Get Position() As Currency
    Position = mPosition
End Property

''
' Sets the current position within the stream.
'
' @param RHS The absolution position within the stream.
' @remarks The position is the index at which the next read or write
' will occur. Position is always zero based, even if a user-supplied
' array is not.
'
Public Property Let Position(ByVal RHS As Currency)
    VerifyOpen
    
    If RHS < 0@ Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Position", RHS)
    mPosition = Int(RHS)
End Property

''
' Closes the MemoryStream and releases any resources allocated.
'
' @remarks If a user-supplied array is being used, it will be released
' at this point. The user then may modify the array as normal.
'
Public Sub CloseStream()
    If mIsClosed Then Exit Sub
    
    If mIsUserArray Then
        UnlockArray
        SAPtr(mBuffer) = 0
    End If
    mIsClosed = True
End Sub

''
' Writes an array of bytes to the underlying stream.
'
' @param Bytes The bytes to be written to the stream.
' @param Offset The startind index in <i>Buffer</i> to begin writing from.
' @param Count The number of bytes to be written to the stream.
' @remarks <p>If the position in the stream is closer to the end than Count bytes, then
' the stream will expand if possible. If the underlying stream is a user-supplied
' byte array, then an exception will occur if the capacity is reached.</p>
' <p>The stream must be open to write to it, otherwise an exception will be thrown.</p>
'
Public Sub WriteBlock(ByRef Bytes() As Byte, ByVal offset As Long, ByVal Count As Long)
    VerifyWritable
    VerifyOpen
    VerifyBuffer
    
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Buffer")
    If offset < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "Offset", offset)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If offset + Count - 1 > UBound(Bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
        
    If Count = 0 Then Exit Sub
    EnsureCapacity mPosition + Count
    CopyMemory mBuffer(mIndex + mPosition), Bytes(offset), Count
    mPosition = mPosition + Count
    If mPosition > mLength Then mLength = mPosition
End Sub

''
' Writes a single byte to the underlying stream.
'
' @param value The byte to be written to the stream.
' @remarks <p>If the position to be written exceeds the current capacity, then
' the stream will be expanded if possible. If the underlying stream is a
' user-supply byte array, then the capacity cannot be modified and an
' exception will be thrown.</p>
' <p>The stream must be open to write to it, otherwise an exception will be thrown.</p>
'
Public Sub WriteByte(ByVal value As Byte)
    VerifyWritable
    VerifyOpen
    VerifyBuffer
    
    mBuffer(mPosition + mIndex) = value
    mPosition = mPosition + 1
    If mPosition > mLength Then mLength = mPosition
End Sub

''
' Reads a specifiec number of bytes from the stream.
'
' @param Bytes The destination for the bytes to be read from the stream.
' @param Offset The startind index in <i>buffer</i> to begin storing the bytes read from the stream.
' @param Count The maximum number of bytes to be read from the stream.
' @return The total number of bytes read from the stream.
' @remarks <p>If Offset plus Count exceeds the end of <i>Buffer</i> then an exception will be thrown.
' If the current position in the stream is closer to the end than Count bytes, then only the
' remaining bytes will be read.</p>
' <p>The stream must be open to read from it, otherwise an exception will be thrown.</p>
' <p>After reading from the stream, the current position will advance the number of byte as returend.</p>
'
Public Function ReadBlock(ByRef Bytes() As Byte, ByVal offset As Long, ByVal Count As Long) As Long
    VerifyOpen
    VerifyBuffer
    
    If cArray.IsNull(Bytes) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Buffer")
    If offset < LBound(Bytes) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "Offset")
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    If offset + Count - 1 > UBound(Bytes) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Count")
    
    If mPosition >= mLength Then Exit Function
    If mPosition + Count > mLength Then Count = mLength - mPosition
    
    CopyMemory Bytes(offset), mBuffer(mPosition + mIndex), Count
    mPosition = mPosition + Count
    ReadBlock = Count
End Function

''
' Reads a single byte from the stream.
'
' @return The byte value from the current position in the stream.
' @remarks <p>The stream must be open for read from it, otherwise an exception will be thrown.</p>
' <p>If the end of the stream is reached during a read, then a value of
' -1 is returned.</p>
' <p>Each read from the stream will advance the position one byte.</p>
'
Public Function ReadByte() As Long
    If mPosition >= mLength Then
        ReadByte = -1
        Exit Function
    End If
    
    ReadByte = mBuffer(mPosition + mIndex)
    mPosition = mPosition + 1
End Function

''
' Changes the position within the stream.
'
' @param Offset The number of bytes to move.
' @param Loc The starting position in the stream to move from.
' @return The actual position moved to.
' @param <p>The offset is applied to a starting position refered to by <i>Loc</i>.
' the <i>Offset</i> can be negative to move the position backwards.</p>
' <p>If the position ends up being before the start of the stream an exception will be thrown.
' The position can be set passed the end of the stream. This will cause the stream to be
' expanded on the next access. If the stream expansion is not supported an exception
' will be thrown. This may be the case when using a user-supplied byte array.</p>
'
Public Function SeekPosition(ByVal offset As Currency, ByVal Loc As SeekOrigin) As Currency
    offset = Int(offset)
    Select Case Loc
        Case FromBeginning
            Position = offset
        Case FromCurrent
            Position = mPosition + offset
        Case FromEnd
            Position = (mLength - 1) + offset
    End Select
    SeekPosition = mPosition
End Function

''
' Sets the current length of the stream.
'
' @param value The size of the stream in bytes.
' @remarks <p>The shortening the length of the stream will truncate the excess data. If
' increasing the length, then the new exposed data will be set to zero.</p>
' <p>If a user-supplied byte array is being used, this function has limited capabilites.
' The size of the stream can be changed to withing the origin size of the user-supplied
' byte array. If a length greater than the original array size is set, then an exception
' is thrown.</p>
'
Public Sub SetLength(ByVal value As Currency)
    Dim i As Long
    
    VerifyWritable
    value = Int(value)
    Select Case value
        Case mLength + 1 To mCapacity
            For i = mLength To value - 1
                mBuffer(mIndex + i) = 0
            Next i
        Case Is > mCapacity
            capacity = value
        Case Is < 0@
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Value", value)
    End Select
    mLength = value
End Sub

''
' Returns a byte array copy of the underlying byte stream.
'
' @return The byte array copy of the underlying stream.
' @remarks If the underlying stream is a user-supplied byte array, then only
' that portion determined by Count when the stream was created is returned, not
' the entire array.
'
Public Function ToArray() As Byte()
    Dim ret() As Byte
    
    VerifyBuffer
    
    ReDim ret(0 To mLength - 1)
    CopyMemory ret(0), mBuffer(mIndex), mLength
    ToArray = ret
End Function

''
' This has no purpose in a MemoryStream
'
Public Sub Flush()
    ' just here for consistency
End Sub

''
' Writes the underlying stream to another stream.
'
' @param Stream The stream to write the underlying stream to.
'
Public Sub WriteTo(ByVal Stream As Stream)
    VerifyBuffer
    VerifyOpen
    
    If Stream Is Nothing Then _
        Throw New ArgumentNullException
    Stream.WriteBlock mBuffer, mIndex, mLength
End Sub


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

''
' Initializes the MemoryStream to a user-supplied byte array.
' This method should never be called directly. Only the constructors
' found in Constructors.cls should call this method.
Friend Sub Init(ByRef Buffer() As Byte, ByRef vIndex As Variant, ByRef vCount As Variant, ByVal Writable As Boolean)
    If cArray.IsNull(Buffer) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Bytes")
    
    mCanWrite = Writable
    mIsUserArray = True
    If IsMissing(vIndex) Then
        mIndex = LBound(Buffer)
    Else
        mIndex = vIndex
    End If
    If IsMissing(vCount) Then
        mCapacity = UBound(Buffer) - mIndex + 1
    Else
        mCapacity = vCount
    End If
    mLength = mCapacity
    
    If mIndex < LBound(Buffer) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_LBound), "vIndex", mIndex)
    If mIndex + mCapacity - 1 > UBound(Buffer) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "vCount")
        
    SAPtr(mBuffer) = SAPtr(Buffer)
    LockArray
    mBufferSet = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LockArray()
    SafeArrayLock cArray.GetValidArrayPointer(mBuffer)
End Sub

Private Sub UnlockArray()
    SafeArrayUnlock cArray.GetValidArrayPointer(mBuffer)
End Sub

' verifies that the internal buffer has been allocated.
Private Sub VerifyBuffer()
    If Not mBufferSet Then
        ReDim mBuffer(0 To mCapacity - 1)
        mCanWrite = True
        mBufferSet = True
    End If
    EnsureCapacity mPosition + 1
End Sub

Private Sub VerifyOpen()
    If mIsClosed Then Throw Cor.NewObjectDisposedException("MemoryStream", "Cannot perform action on a closed stream.")
End Sub

Private Sub VerifyWritable()
    If Not mCanWrite Then Throw Cor.NewNotSupportedException("Stream is ReadOnly.")
End Sub

' Ensures that the required capacity is met. If the capacity needs to
' be increased, an exception may be thrown by the Capacity property when set.
Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    Dim NewCapacity As Long
    
    If RequiredCapacity <= mCapacity Then Exit Sub
    NewCapacity = mCapacity * 2
    If NewCapacity < RequiredCapacity Then NewCapacity = RequiredCapacity + DEF_CAPACITY
    capacity = NewCapacity
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mCapacity = DEF_CAPACITY
    mCanWrite = True
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mBuffer = .ReadProperty("Buffer")
        mCanWrite = .ReadProperty("CanWrite")
        mPosition = .ReadProperty("Position")
        mCapacity = .ReadProperty("Capacity")
        mLength = .ReadProperty("Length")
        mIsClosed = .ReadProperty("IsClosed")
        mIndex = .ReadProperty("Index")
    End With
End Sub

Private Sub Class_Terminate()
    CloseStream
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Buffer", mBuffer
        .WriteProperty "CanWrite", mCanWrite
        .WriteProperty "Position", mPosition
        .WriteProperty "Capacity", mCapacity
        .WriteProperty "Length", mLength
        .WriteProperty "IsClosed", mIsClosed
        .WriteProperty "Index", mIndex
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
'   Stream Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Stream_BeginRead(Buffer() As Byte, ByVal offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Set Stream_BeginRead = BeginRead(Buffer, offset, Count, Callback, State)
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Set Stream_BeginWrite = BeginWrite(Buffer, offset, Count, Callback, State)
End Function

Private Property Get Stream_CanRead() As Boolean
    Stream_CanRead = CanRead
End Property

Private Property Get Stream_CanSeek() As Boolean
    Stream_CanSeek = CanSeek
End Property

Private Property Get Stream_CanWrite() As Boolean
    Stream_CanWrite = CanWrite
End Property

Private Sub Stream_CloseStream()
    CloseStream
End Sub

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long
    Stream_EndRead = EndRead(AsyncResult)
End Function

Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)
    EndWrite AsyncResult
End Sub

Private Function Stream_Equals(value As Variant) As Boolean
    Stream_Equals = Equals(value)
End Function

Private Sub Stream_Flush()
    Flush
End Sub

Private Function Stream_GetHashCode() As Long
    Stream_GetHashCode = GetHashCode
End Function

Private Property Get Stream_Length() As Currency
    Stream_Length = Length
End Property

Private Property Let Stream_Position(ByVal RHS As Currency)
    Position = RHS
End Property

Private Property Get Stream_Position() As Currency
    Stream_Position = Position
End Property

Private Function Stream_ReadBlock(Buffer() As Byte, ByVal offset As Long, ByVal Count As Long) As Long
    Stream_ReadBlock = ReadBlock(Buffer, offset, Count)
End Function

Private Function Stream_ReadByte() As Long
    Stream_ReadByte = ReadByte
End Function

Private Function Stream_SeekPosition(ByVal offset As Currency, ByVal Origin As SeekOrigin) As Currency
    Stream_SeekPosition = SeekPosition(offset, Origin)
End Function

Private Sub Stream_SetLength(ByVal value As Currency)
    SetLength value
End Sub

Private Function Stream_ToString() As String
    Stream_ToString = ToString
End Function

Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal offset As Long, ByVal Count As Long)
    WriteBlock Buffer, offset, Count
End Sub

Private Sub Stream_WriteByte(ByVal value As Byte)
    WriteByte value
End Sub
