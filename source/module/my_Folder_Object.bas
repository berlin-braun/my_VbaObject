Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private my_fso  As Object ' Scripting.FileSystemObject
Private my_fdr  As Object
'


Private Sub Class_Terminate()

  Set my_fso = Nothing
  Set my_fdr = Nothing
  
End Sub


' Properties - Start

Public Property Get Attributes() As Long
  Attributes = my_fdr.Attributes
End Property

Public Property Get DateCreated() As Date
  DateCreated = my_fdr.DateCreated
End Property

Public Property Get DateLastAccessed() As Date
  DateLastAccessed = my_fdr.DateLastAccessed
End Property

Public Property Get DateLastModified() As Date
  DateLastModified = my_fdr.DateLastModified
End Property

Public Property Get Drive() As Object
  Set Drive = my_fdr.Drive
End Property

Public Property Get IsRootFolder() As Boolean
  IsRootFolder = my_fdr.IsRootFolder
End Property

Public Property Let Name(ByVal str_Folder As String)
  
  Set my_fso = Nothing
  Set my_fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
  
  Set my_fdr = my_fso.GetFolder(str_Folder)
  
End Property

Public Property Get Name() As String
  Name = my_fdr.Name
End Property

Public Property Get ParentFolder() As String
  ParentFolder = my_fdr.ParentFolder
End Property

Public Property Get Path() As String
  Path = my_fdr.Path
End Property

Public Property Get ShortName() As String
  ShortName = my_fdr.ShortName
End Property

Public Property Get ShortPath() As String
  ShortPath = my_fdr.ShortPath
End Property

Public Property Get Typ() As String
  Typ = my_fdr.Type
End Property

' Properties - End


' Methods - Start

Public Function Delete(Optional force As Boolean = False)
  
  my_fso.DeleteFolder my_fdr.Path, force
  
End Function

Public Function Move(ByVal destination As String)
  
  my_fso.MoveFolder my_fdr.Path, destination
  
End Function

Public Function Copy(ByVal destination As String)
  
  my_fso.CopyFolder my_fdr.Path, destination
  
End Function

' Methods - End