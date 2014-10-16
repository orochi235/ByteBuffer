VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ByteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_buffer() As Byte
Private m_capacity As Long      ' The capacity of the buffer.
Private m_length As Long        ' The number of bytes currently stored in the buffer.

Private Const DEFAULT_LENGTH = 10000

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Property Get ByteAt(ByVal Index As Long) As Byte
' Returns the byte at the specified position.
    ByteAt = m_buffer(Index)
End Property

Public Property Let ByteAt(ByVal Index As Long, ByVal Value As Byte)
    m_buffer(Index) = Value
End Property

'..................................................................................................................

Public Property Get Data() As Byte()
Attribute Data.VB_UserMemId = 0
    Data = m_buffer
End Property

'..................................................................................................................

Public Property Get Capacity() As Long
    Capacity = m_capacity
End Property

Public Property Let Capacity(ByVal Size As Long)
    Resize Size
End Property

'..................................................................................................................

Public Property Get Length() As Long
    Length = m_length
End Property

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Class_Initialize()
    m_length = 0
    Initialize DEFAULT_LENGTH
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Sub AppendByte(ByVal Data As Byte)
' Appends a byte to the buffer.
    
    If m_length + 1 > m_capacity Then Extend m_length * 2   ' Double the size of the buffer; this will prevent frequent, costly resizing operations.
    
    Me.ByteAt(m_length) = Data
    m_length = m_length + 1
    
End Sub

Public Sub AppendBytes(ByRef Bytes() As Byte)
' Appends a byte array to the buffer.
    
    Dim CurrentPosition As Long
    CurrentPosition = m_length
    
    If m_length + (UBound(Bytes) - 1) > m_capacity Then Extend m_length + (UBound(Bytes) - 1)
    
    Dim i As Long
    For i = 0 To UBound(Bytes)
        Me.ByteAt(CurrentPosition) = Bytes(i)
        CurrentPosition = CurrentPosition + 1
    Next i
    
    m_length = m_length + UBound(Bytes) + 1                 ' Push the length pointer ahead by the number of bytes that have been added.

End Sub

Public Sub AppendChunk(ByRef Chunk As ByteBuffer)
' Appends a chunk of binary data from another ByteBuffer object.
    
    If m_capacity < m_length + Chunk.Length Then
        Resize m_length + Chunk.Length
    End If
    
    Dim Index As Long
    For Index = 0 To Chunk.Length - 1
        Me.ByteAt(m_length + Index) = Chunk.ByteAt(Index)
    Next Index
    
    m_length = m_length + Chunk.Length
    
End Sub

Public Sub AppendString(ByVal Data As String)
' Appends a chunk of binary data represented in a string.
' Although all VB strings are MBCS, we're only interested in the ASCII equivalent of each character.
    
    Dim CurrentPosition As Long
    CurrentPosition = m_length
    
    If m_length + Len(Data) > m_capacity Then Extend m_length + Len(Data)
    
    Dim i As Long
    For i = 0 To Len(Data) - 1
        Me.ByteAt(CurrentPosition) = Asc(Mid(Data, i + 1, 1))
        CurrentPosition = CurrentPosition + 1
    Next i
    
    m_length = m_length + Len(Data)                         ' Push the length pointer ahead by the number of bytes that have been added.
    
End Sub

Public Function Copy(Optional ByVal MaxLength As Long = 0) As ByteBuffer
' Creates a new instance of this object and copies the data.
' If MaxLength > 0, the copy will only contain the specified number of bytes.
    
    Dim CopyLength As Long
    If MaxLength < 1 Then
        CopyLength = Me.Length
    Else
        CopyLength = MaxLength
    End If

    Dim OutBuffer As New ByteBuffer
    OutBuffer.Initialize IIf(MaxLength = 0, Me.Capacity, MaxLength)
    
    Dim i As Long
    For i = 0 To CopyLength - 1
        OutBuffer.AppendByte Me.ByteAt(i)
    Next i
    
    Set Copy = OutBuffer
    
End Function

Public Sub Extend(ByVal Capacity As Long)
' Extends the buffer to a specified capacity.  If the new capacity is smaller than the old one, instead returns an error.
    
    If Capacity >= m_capacity Then
        ReDim Preserve m_buffer(Capacity - 1)
        m_capacity = Capacity
    Else
        Err.Raise vbObjectError, "ByteBuffer", "New length is smaller than old length."
    End If
    
End Sub

Public Sub Initialize(ByVal Capacity As Long)
    ReDim m_buffer(Capacity - 1)
    m_capacity = Capacity
End Sub

Public Sub Load(ByVal FileName As String)
' Loads data from the specified file.
    
    Dim FileNumber As Integer
    FileNumber = FreeFile()
    
    Dim FileLength As Long
    FileLength = FileLen(FileName)
    
    Initialize FileLength                                   ' Resize the byte array to the proper length.

    Open FileName For Binary As #FileNumber
    Get #FileNumber, 1, m_buffer
    Close FileNumber

    m_length = FileLength

End Sub

Public Sub Resize(ByVal Capacity As Long)
' Resizes the buffer (in either direction), preserving its current contents.
    ReDim Preserve m_buffer(Capacity - 1)
    m_capacity = Capacity
    
    If m_length > m_capacity Then m_length = m_capacity
    
End Sub

Public Sub Save(ByVal FileName As String)
' Saves the contents of the buffer to the specified file.

    Dim MyCopy As ByteBuffer
    Set MyCopy = Me.Copy
    MyCopy.Truncate
    MyCopy.WriteToFile FileName
    Set MyCopy = Nothing

End Sub

Public Sub SaveAsText(ByVal FileName As String)
' Saves the contents of the buffer to a specified file as text.

    ' TODO: Write this.

End Sub

Friend Sub Truncate()
' Removes all trailing spaces from the array.  Should be done prior to file output.
    ReDim Preserve m_buffer(m_length - 1)
End Sub

Friend Sub WriteToFile(ByVal FileName As String)
' Writes the contents of the buffer to a file.  Should only be invoked by Save(), which copies a buffer, truncates it, and then calls this.

    ' Delete any existing file.
    On Error Resume Next
    Kill FileName
    On Error GoTo 0

    Dim FileNumber As Integer
    FileNumber = FreeFile

    Open FileName For Binary As #FileNumber
    Put #FileNumber, 1, m_buffer
    Close #FileNumber

End Sub
