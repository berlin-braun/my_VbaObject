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


' ----------------------------------------------------------------
' Procedure Name:   textstream_Write_Text
' Purpose:          Writes a specified string to a TextStream file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter filename (String): the specific filename
' Parameter text (String): Required. The text you want to write to the file.
' ----------------------------------------------------------------
Public Function textstream_Write_Text(ByVal filename As String _
                                    , ByVal text As String)
  
  Dim my_Text As New my_TextStream_Object
  
  Set my_Text = textstream_INIT(filename, ForAppending, , , True)
  
  my_Text.Write_Text text
  my_Text.Close_Stream
  
  Set my_Text = Nothing
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   textstream_WriteLine
' Purpose:          Writes a specified string and newline character to a TextStream file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter filename (String): the specific filename
' Parameter text (String): Optional. The text you want to write to the file. If omitted, a newline character is written to the file.
' ----------------------------------------------------------------
Public Function textstream_WriteLine(ByVal filename As String _
                                    , ByVal text As String)
  
  Dim my_Text As New my_TextStream_Object
  
  Set my_Text = textstream_INIT(filename, ForAppending, True, , True)
  
  my_Text.WriteLine text
  my_Text.Close_Stream
  
  Set my_Text = Nothing
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   textstream_WriteBlankLines
' Purpose:          Writes a specified number of newline characters to a TextStream file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter filename (String): the specific filename
' Parameter lines (Long): Required. Number of newline characters you want to write to the file.
' ----------------------------------------------------------------
Public Function textstream_WriteBlankLines(ByVal filename As String _
                                         , ByVal lines As Long)
  
  Dim my_Text As New my_TextStream_Object
  
  Set my_Text = textstream_INIT(filename, ForAppending, True, , True)
  
  my_Text.WriteBlankLines lines
  my_Text.Close_Stream
  
  Set my_Text = Nothing
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   textstream_ReadAll
' Purpose:          Reads an entire TextStream file and returns the resulting string.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter filename (String): the specific filename
' ----------------------------------------------------------------
Public Function textstream_ReadAll(ByVal filename As String) As String
  
  Dim my_Text     As New my_TextStream_Object
  Dim str_Ret     As String
  
  Set my_Text = textstream_INIT(filename, ForReading, False, , False)
  
  str_Ret = my_Text.ReadAll
  
  Set my_Text = Nothing
  
  textstream_ReadAll = str_Ret
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   textstream_ReadLine
' Purpose:          Reads an entire line (up to, but not including, the newline character) from a TextStream file and returns the resulting string.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter filename (String): the specific filename
' ----------------------------------------------------------------
Public Function textstream_ReadLine(ByVal filename As String) As String
  
  Dim my_Text     As New my_TextStream_Object
  Dim str_Ret     As String
  
  Set my_Text = textstream_INIT(filename, ForReading, False, , False)
  
  str_Ret = my_Text.ReadLine
  
  Set my_Text = Nothing
  
  textstream_ReadLine = str_Ret
  
End Function