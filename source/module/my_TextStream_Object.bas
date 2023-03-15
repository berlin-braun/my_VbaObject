Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private my_fso  As Scripting.FileSystemObject
Private my_txs  As Object
'


Private Sub Class_Terminate()

  Set my_fso = Nothing
  Set my_txs = Nothing
  
End Sub


' Properties - Start

Public Property Get AtEndOfLine() As Boolean
  AtEndOfLine = my_txs.AtEndOfLine
End Property

Public Property Get AtEndOfStream() As Boolean
  AtEndOfStream = my_txs.AtEndOfStream
End Property

Public Property Get Column() As Long
  Column = my_txs.Column
End Property

Public Property Get Line() As String
  Line = my_txs.Line
End Property

' Properties - End


' Methods - Start

Public Function Close_Stream()
  my_txs.Close
End Function

Public Function Read(ByVal characters As Long)
  Read = my_txs.Read(characters)
End Function

Public Function ReadAll()
  ReadAll = my_txs.ReadAll
End Function

Public Function ReadLine()
  ReadLine = my_txs.ReadLine
End Function

Public Function Skip(ByVal characters As Long)
  my_txs.Skip characters
End Function

Public Function SkipLine()
  my_txs.SkipLine
End Function

Public Function Write_Text(ByVal text As String)
  my_txs.Write text
End Function

Public Function WriteBlankLines(ByVal lines As Long)
  my_txs.WriteBlankLines lines
End Function

Public Function WriteLine(ByVal text As String)
  my_txs.WriteLine text
End Function

Public Function CreateTextFile(ByVal filename As String _
                    , Optional ByVal overwrite As Boolean = True _
                    , Optional ByVal unicode As Boolean = False) As Object
  
  Set my_fso = Nothing
  Set my_txs = Nothing
  
  Set my_fso = New Scripting.FileSystemObject
  Set my_txs = my_fso.CreateTextFile(filename, overwrite, unicode)
  
  Set CreateTextFile = my_txs
  
End Function

Public Function OpenTextFile(ByVal filename As String _
                  , Optional ByVal mode As IOMode = ForAppending _
                  , Optional ByVal create As Boolean = False _
                  , Optional ByVal format As Tristate = TristateUseDefault) As Object
  
  Set my_fso = Nothing
  Set my_txs = Nothing
  
  Set my_fso = New Scripting.FileSystemObject
  Set my_txs = my_fso.OpenTextFile(filename, mode, create, format)
  
  Set OpenTextFile = my_txs
  
End Function

' Methods - End