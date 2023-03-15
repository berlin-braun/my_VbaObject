Option Compare Database
Option Explicit
'
'
' Factory: static properties and methods for class my_FileSystemObject_Object
'
'



Public Function Drives() As Object
  Dim m_FSO As New my_FileSystemObject_Object
  
  Set Drives = m_FSO.Drives
  
  Set m_FSO = Nothing
End Function


Public Function filesystemobject_BuildPath(ByVal Path As String _
                                         , ByVal Name As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_BuildPath = m_FSO.BuildPath(Path, Name)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_CopyFile(ByVal source As String _
                                        , ByVal destination As String _
                               , Optional ByVal overwrite As Boolean = True)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.CopyFile source, destination, overwrite
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_CopyFolder(ByVal source As String _
                                          , ByVal destination As String _
                                 , Optional ByVal overwrite As Boolean = True)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.CopyFolder source, destination, overwrite
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_CreateFolder(ByVal foldername As String)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.CreateFolder foldername
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_CreateTextFile(ByVal filename As String _
                                     , Optional ByVal overwrite As Boolean = True _
                                     , Optional ByVal unicode As Boolean = False)
  Dim m_FSO As New my_FileSystemObject_Object

  m_FSO.CreateTextFile filename, overwrite, unicode
  
  Set m_FSO = Nothing
End Function


Public Function filesystemobject_DeleteFile(ByVal filespec As String _
                                 , Optional ByVal force As Boolean = False)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.DeleteFile filespec, force
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_DeleteFolder(ByVal folderspec As String _
                                   , Optional ByVal force As Boolean = False)
  Dim m_FSO As New my_FileSystemObject_Object

  m_FSO.DeleteFolder folderspec, force
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_DriveExists(ByVal drivespec As String) As Boolean
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_DriveExists = m_FSO.DriveExists(drivespec)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_FileExists(ByVal filespec As String) As Boolean
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_FileExists = m_FSO.FileExists(filespec)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_FolderExists(ByVal folderspec As String) As Boolean
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_FolderExists = m_FSO.FolderExists(folderspec)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_GetAbsolutePathName(ByVal pathspec As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetAbsolutePathName = m_FSO.GetAbsolutePathName(pathspec)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_GetBaseName(ByVal Path As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetBaseName = m_FSO.GetBaseName(Path)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_GetDrive(ByVal drivespec As String) As my_Drive_Object
  Dim m_FSO As New my_FileSystemObject_Object
  
  Set filesystemobject_GetDrive = m_FSO.GetDrive(drivespec)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_GetDriveName(ByVal Path As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetDriveName = m_FSO.GetDriveName(Path)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_GetExtensionName(ByVal Path As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetExtensionName = m_FSO.GetExtensionName(Path)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_GetFile(ByVal filespec As String) As my_File_Object
  Dim m_FSO As New my_FileSystemObject_Object
  
  Set filesystemobject_GetFile = m_FSO.GetFile(filespec)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_GetFileName(ByVal pathspec As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetFileName = m_FSO.GetFileName(pathspec)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_GetFolder(ByVal folderspec As String) As my_Folder_Object
  Dim m_FSO As New my_FileSystemObject_Object
    
  Set filesystemobject_GetFolder = m_FSO.GetFolder(folderspec)
  
End Function

Public Function filesystemobject_GetParentFolderName(ByVal Path As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetParentFolderName = m_FSO.GetParentFolderName(Path)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_GetSpecialFolder(ByVal folderspec As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetSpecialFolder = m_FSO.GetSpecialFolder(folderspec)
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_GetTempName() As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetTempName = m_FSO.GetTempName
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_MoveFile(ByVal source As String _
                                        , ByVal destination As String)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.MoveFile source, destination
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_MoveFolder(ByVal source As String _
                                          , ByVal destination As String)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.MoveFolder source, destination
  
  Set m_FSO = Nothing
End Function

Public Function filesystemobject_OpenTextFile(ByVal filename As String _
                                   , Optional ByVal mode As IOMode = ForAppending _
                                   , Optional ByVal create As Boolean = False _
                                   , Optional ByVal format As Tristate = TristateUseDefault) As my_File_Object
  Dim m_FSO As New my_FileSystemObject_Object
  
  Set filesystemobject_OpenTextFile = m_FSO.OpenTextFile(filename, mode, create, format)

  Set m_FSO = Nothing
End Function