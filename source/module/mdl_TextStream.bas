Option Compare Database
Option Explicit
'
'
' Factory: static functions for class my_TextStream_Object
'
'

Public Function textstream_INIT(ByVal filename As String _
                     , Optional ByVal mode As my_IOMode = ForAppending _
                     , Optional ByVal create As Boolean = False _
                     , Optional ByVal format As my_Tristate = TristateUseDefault _
                     , Optional ByVal overwrite As Boolean = True _
                     , Optional ByVal unicode As Boolean = True) As my_TextStream_Object
  
  Dim m_Stream As New my_TextStream_Object
  
  If filesystemobject_FileExists(filename) Then
    m_Stream.OpenTextFile filename, mode, create, format
  Else
    m_Stream.CreateTextFile filename, overwrite, unicode
  End If
  
  Set textstream_INIT = m_Stream
  
  Set m_Stream = Nothing
  
End Function


Public Function textstream_Write_Text(ByVal filename As String _
                                    , ByVal text As String)
  
  Dim my_Text As New my_TextStream_Object
  
  Set my_Text = textstream_INIT(filename, ForAppending, , , True)
  
  my_Text.Write_Text text
  my_Text.Close_Stream
  
  Set my_Text = Nothing
  
End Function

Public Function textstream_WriteLine(ByVal filename As String _
                                    , ByVal text As String)
  
  Dim my_Text As New my_TextStream_Object
  
  Set my_Text = textstream_INIT(filename, ForAppending, True, , True)
  
  my_Text.WriteLine text
  my_Text.Close_Stream
  
  Set my_Text = Nothing
  
End Function

Public Function textstream_WriteBlankLines(ByVal filename As String _
                                         , ByVal lines As Long)
  
  Dim my_Text As New my_TextStream_Object
  
  Set my_Text = textstream_INIT(filename, ForAppending, True, , True)
  
  my_Text.WriteBlankLines lines
  my_Text.Close_Stream
  
  Set my_Text = Nothing
  
End Function

Public Function textstream_ReadAll(ByVal filename As String) As String
  
  Dim my_Text     As New my_TextStream_Object
  Dim str_Ret     As String
  
  Set my_Text = textstream_INIT(filename, ForReading, False, , False)
  
  str_Ret = my_Text.ReadAll
  
  Set my_Text = Nothing
  
  textstream_ReadAll = str_Ret
  
End Function

Public Function textstream_ReadLine(ByVal filename As String) As String
  
  Dim my_Text     As New my_TextStream_Object
  Dim str_Ret     As String
  
  Set my_Text = textstream_INIT(filename, ForReading, False, , False)
  
  str_Ret = my_Text.ReadLine
  
  Set my_Text = Nothing
  
  textstream_ReadLine = str_Ret
  
End Function