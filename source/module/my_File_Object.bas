Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private my_fso As Scripting.FileSystemObject
Private my_fle As Object
'


Private Sub Class_Terminate()
  
  Set my_fso = Nothing
  Set my_fle = Nothing
  
End Sub


' Properties - Start

Public Property Get Attributes() As Long
  Attributes = my_fle.Attributes
End Property

Public Property Get DateCreated() As Date
  DateCreated = my_fle.DateCreated
End Property

Public Property Get DateLastAccessed() As Date
  DateLastAccessed = my_fle.DateLastAccessed
End Property

Public Property Get DateLastModified() As Date
  DateLastModified = my_fle.DateLastModified
End Property

Public Property Get Drive() As Object
  
  Set Drive = my_fle.Drive

End Property

Public Property Let Name(ByVal str_File As String)
  
  Set my_fso = Nothing
  Set my_fso = New Scripting.FileSystemObject
  
  Set my_fle = my_fso.GetFile(str_File)
  
End Property

Public Property Get Name() As String
  Name = my_fle.Name
End Property

Public Property Get ParentFolder() As Object
  
  Set ParentFolder = my_fle.ParentFolder

End Property

Public Property Get Path() As String
  Path = my_fle.Path
End Property

Public Property Get ShortName() As String
  ShortName = my_fle.ShortName
End Property

Public Property Get ShortPath() As String
  ShortPath = my_fle.ShortPath
End Property

Public Property Get Size() As Long
  Size = my_fle.Size
End Property

Public Property Get Typ() As String
  Typ = my_fle.Type
End Property


' Properties - End


' Methods - Start

Public Function Delete(Optional force As Boolean = False)
  
  my_fso.DeleteFile my_fle.Path
  
End Function

Public Function Move(ByVal destination As String)
  
  my_fso.MoveFile my_fle.Path, destination
  
End Function

Public Function Copy(ByVal destination As String)
  
  my_fso.CopyFile my_fle.Path, destination
  
End Function

Public Function CreateTextFile(filename As String _
                    , Optional overwrite As Boolean = True _
                    , Optional unicode As Boolean = False)
  Set my_fso = Nothing
  Set my_fle = Nothing
  
  Set my_fso = New Scripting.FileSystemObject
  Set my_fle = my_fso.CreateTextFile(filename, overwrite, unicode)
  
End Function

Public Function OpenAsTextStream(ByVal filename As String _
                      , Optional ByVal mode As IOMode = ForAppending _
                      , Optional ByVal format As Tristate = TristateUseDefault) As Object
  Dim my_txs As Object

  Set my_fso = Nothing
  Set my_txs = Nothing

  Set my_fso = New Scripting.FileSystemObject
  Set my_fle = my_fso.GetFile(filename)
  Set my_txs = my_fle.OpenAsTextStream(filename, mode, format)

  Set OpenAsTextStream = my_txs

  Set my_txs = Nothing
  
End Function

' Methods - End