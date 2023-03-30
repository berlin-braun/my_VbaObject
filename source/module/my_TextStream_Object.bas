Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private my_fso  As Object ' Scripting.FileSystemObject
Private my_txs  As Object

Public Enum my_Tristate
  TristateFalse = 0
  TristateMixed = -2
  TristateTrue = -1
  TristateUseDefault = -2
End Enum

Public Enum my_IOMode
  ForAppending = 8
  ForReading = 1
  ForWriting = 2
End Enum
'


Private Sub Class_Terminate()

  Set my_fso = Nothing
  Set my_txs = Nothing
  
End Sub


' Properties - Start


' ----------------------------------------------------------------
' Procedure Name:   AtEndOfLine
' Purpose:          Read-only property that returns True if the file pointer immediately precedes the end-of-line marker in a TextStream file; False if it does not.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Boolean
' ----------------------------------------------------------------
Public Property Get AtEndOfLine() As Boolean
  AtEndOfLine = my_txs.AtEndOfLine
End Property


' ----------------------------------------------------------------
' Procedure Name:   AtEndOfStream
' Purpose:          Read-only property that returns True if the file pointer is at the end of a TextStream file; False if it is not.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Boolean
' ----------------------------------------------------------------
Public Property Get AtEndOfStream() As Boolean
  AtEndOfStream = my_txs.AtEndOfStream
End Property


' ----------------------------------------------------------------
' Procedure Name:   Column
' Purpose:          Read-only property that returns the column number of the current character position in a TextStream file.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' ----------------------------------------------------------------
Public Property Get Column() As Long
  Column = my_txs.Column
End Property


' ----------------------------------------------------------------
' Procedure Name:   Line
' Purpose:          Read-only property that returns the current line number in a TextStream file.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get Line() As String
  Line = my_txs.Line
End Property

' Properties - End


' Methods - Start


' ----------------------------------------------------------------
' Procedure Name:   Close_Stream
' Purpose:          Closes an open TextStream file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' ----------------------------------------------------------------
Public Function Close_Stream()
  my_txs.Close
End Function


' ----------------------------------------------------------------
' Procedure Name:   Read
' Purpose:          Reads a specified number of characters from a TextStream file and returns the resulting string.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter characters (Long): Required. Number of characters that you want to read from the file.
' ----------------------------------------------------------------
Public Function Read(ByVal characters As Long)
  Read = my_txs.Read(characters)
End Function


' ----------------------------------------------------------------
' Procedure Name:   ReadAll
' Purpose:          Reads an entire TextStream file and returns the resulting string.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' ----------------------------------------------------------------
Public Function ReadAll()
  ReadAll = my_txs.ReadAll
End Function


' ----------------------------------------------------------------
' Procedure Name:   ReadLine
' Purpose:          Reads an entire line (up to, but not including, the newline character) from a TextStream file and returns the resulting string.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' ----------------------------------------------------------------
Public Function ReadLine()
  ReadLine = my_txs.ReadLine
End Function


' ----------------------------------------------------------------
' Procedure Name:   Skip
' Purpose:          Skips a specified number of characters when reading a TextStream file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter characters (Long): Required. Number of characters to skip when reading a file.
' ----------------------------------------------------------------
Public Function Skip(ByVal characters As Long)
  my_txs.Skip characters
End Function


' ----------------------------------------------------------------
' Procedure Name:   SkipLine
' Purpose:          Skips the next line when reading a TextStream file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' ----------------------------------------------------------------
Public Function SkipLine()
  my_txs.SkipLine
End Function


' ----------------------------------------------------------------
' Procedure Name:   Write_Text
' Purpose:          Writes a specified string to a TextStream file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter text (String): Required. The text you want to write to the file.
' ----------------------------------------------------------------
Public Function Write_Text(ByVal text As String)
  my_txs.Write text
End Function


' ----------------------------------------------------------------
' Procedure Name:   WriteBlankLines
' Purpose:          Writes a specified number of newline characters to a TextStream file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter lines (Long): Required. Number of newline characters you want to write to the file.
' ----------------------------------------------------------------
Public Function WriteBlankLines(ByVal lines As Long)
  my_txs.WriteBlankLines lines
End Function


' ----------------------------------------------------------------
' Procedure Name:   WriteLine
' Purpose:          Writes a specified string and newline character to a TextStream file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter text (String): Optional. The text you want to write to the file. If omitted, a newline character is written to the file.
' ----------------------------------------------------------------
Public Function WriteLine(ByVal text As String)
  my_txs.WriteLine text
End Function


' ----------------------------------------------------------------
' Procedure Name:   CreateTextFile
' Purpose:          Creates a specified file name and returns a TextStream object that can be used to read from or write to the file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' Parameter filename (String): Required. String expression that identifies the file to create.
' Parameter overwrite (Boolean): Optional. Boolean value that indicates if an existing file can be overwritten. The value is True if the file can be overwritten; False if it can't be overwritten. If omitted, existing files can be overwritten.
' Parameter unicode (Boolean): Optional. Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is True if the file is created as a Unicode file; False if it's created as an ASCII file. If omitted, an ASCII file is assumed.
' ----------------------------------------------------------------
Public Function CreateTextFile(ByVal filename As String _
                    , Optional ByVal overwrite As Boolean = True _
                    , Optional ByVal unicode As Boolean = False) As Object
  
  Set my_fso = Nothing
  Set my_txs = Nothing
  
  Set my_fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
  Set my_txs = my_fso.CreateTextFile(filename, overwrite, unicode)
  
  Set CreateTextFile = my_txs
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   OpenTextFile
' Purpose:          Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' Parameter filename (String): Required. String expression that identifies the file to open.
' Parameter mode (my_IOMode):   Optional. Indicates input/output mode. Can be one of three constants: ForReading, ForWriting, or ForAppending.
' Parameter create (Boolean): Optional. Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created; False if it isn't created. The default is False.
' Parameter format (my_Tristate): Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.
' ----------------------------------------------------------------
Public Function OpenTextFile(ByVal filename As String _
                  , Optional ByVal mode As my_IOMode = ForAppending _
                  , Optional ByVal create As Boolean = False _
                  , Optional ByVal format As my_Tristate = TristateUseDefault) As Object
  
  Set my_fso = Nothing
  Set my_txs = Nothing
  
  Set my_fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
  Set my_txs = my_fso.OpenTextFile(filename, mode, create, format)
  
  Set OpenTextFile = my_txs
  
End Function

' Methods - End