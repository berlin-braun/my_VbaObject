Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private my_fso As Scripting.FileSystemObject
'


Private Sub Class_Initialize()
  
  If my_fso Is Nothing Then
    Set my_fso = New Scripting.FileSystemObject
  End If
  
End Sub

Private Sub Class_Terminate()

  If my_fso Is Nothing = False Then
    Set my_fso = Nothing
  End If
  
End Sub


' Properties - Start

Public Property Get Drives() As Object
  
  Set Drives = my_fso.Drives
  
End Property

' Properties - End


' Methods - Start

Public Function BuildPath(ByVal Path As String _
                        , ByVal Name As String) As String
    
  BuildPath = my_fso.BuildPath(Path, Name)
  
End Function

Public Function CopyFile(ByVal source As String _
                       , ByVal destination As String _
              , Optional ByVal overwrite As Boolean = True)
  
  my_fso.CopyFile source, destination, overwrite
  
End Function

Public Function CopyFolder(ByVal source As String _
                         , ByVal destination As String _
                , Optional ByVal overwrite As Boolean = True)
  
  my_fso.CopyFolder source, destination, overwrite
  
End Function

Public Function CreateFolder(ByVal foldername As String)
  
  my_fso.CreateFolder (foldername)
  
End Function

Public Function CreateTextFile(ByVal filename As String _
                    , Optional ByVal overwrite As Boolean = True _
                    , Optional ByVal unicode As Boolean = False)
  
  Set my_fso = Nothing
  Set my_fso = New Scripting.FileSystemObject
  my_fso.CreateTextFile filename, overwrite, unicode
  
End Function

Public Function DeleteFile(ByVal filespec As String _
                , Optional ByVal force As Boolean = False)
  
  my_fso.DeleteFile filespec, force
  
End Function

Public Function DeleteFolder(ByVal folderspec As String _
                  , Optional ByVal force As Boolean = False)
  
  my_fso.DeleteFolder folderspec, force
  
End Function

Public Function DriveExists(ByVal drivespec As String) As Boolean
  
  DriveExists = my_fso.DriveExists(drivespec)
  
End Function

Public Function FileExists(ByVal filespec As String) As Boolean
  
  FileExists = my_fso.FileExists(filespec)
  
End Function

Public Function FolderExists(ByVal folderspec As String) As Boolean
  
  FolderExists = my_fso.FolderExists(folderspec)
  
End Function

Public Function GetAbsolutePathName(ByVal pathspec As String) As String
  
  GetAbsolutePathName = my_fso.GetAbsolutePathName(pathspec)
  
End Function

Public Function GetBaseName(ByVal Path As String) As String
  
  GetBaseName = my_fso.GetBaseName(Path)
  
End Function

Public Function GetDrive(ByVal drivespec As String) As my_Drive_Object
  
'  Set GetDrive = my_fso.GetDrive(drivespec)
  Set GetDrive = drive_INIT(drivespec)
  
End Function

Public Function GetDriveName(ByVal Path As String) As String
  
  GetDriveName = my_fso.GetDriveName(Path)
  
End Function

Public Function GetExtensionName(ByVal Path As String) As String
  
  GetExtensionName = my_fso.GetExtensionName(Path)
  
End Function

Public Function GetFile(ByVal filespec As String) As my_File_Object
  
'  Set GetFile = my_fso.GetFile(filespec)
  Set GetFile = file_INIT(filespec)
  
End Function

Public Function GetFileName(ByVal pathspec As String) As String
  
  GetFileName = my_fso.GetFileName(pathspec)
  
End Function

Public Function GetFolder(ByVal folderspec As String) As my_Folder_Object
  
'  Set GetFolder = my_fso.GetFolder(folderspec)
  Set GetFolder = folder_INIT(folderspec)
  
End Function

Public Function GetParentFolderName(ByVal Path As String) As String
  
  GetParentFolderName = my_fso.GetParentFolderName(Path)
  
End Function

Public Function GetSpecialFolder(ByVal folderspec As String) As String
  
  GetSpecialFolder = my_fso.GetSpecialFolder(folderspec)
  
End Function

Public Function GetTempName() As String
  
  GetTempName = my_fso.GetTempName
  
End Function

Public Function MoveFile(ByVal source As String _
                       , ByVal destination As String)
  
  my_fso.MoveFile source, destination
  
End Function

Public Function MoveFolder(ByVal source As String _
                         , ByVal destination As String)
  
  my_fso.MoveFolder source, destination
  
End Function

Public Function OpenTextFile(ByVal filename As String _
                  , Optional ByVal mode As IOMode = ForAppending _
                  , Optional ByVal create As Boolean = False _
                  , Optional ByVal format As Tristate = TristateUseDefault) As my_File_Object
  
'  Dim my_txs As Object
  
'  Set my_fso = Nothing
'  Set my_txs = Nothing
'  Set my_fso = New Scripting.FileSystemObject
'  Set my_txs = my_fso.OpenTextFile(filename, mode, create, format)
  
'  Set OpenTextFile = my_txs
'  Set my_txs = Nothing

  Set OpenTextFile = textstream_INIT(filename, mode, create, format)

End Function

' Methods - End