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


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_BuildPath
' Purpose:          Combines a folder path and the name of a folder or file and returns the combination with valid path separators.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter Path (String): Required. Existing path with which name is combined. Path can be absolute or relative and need not specify an existing folder.
' Parameter Name (String): Required. Name of a folder or file being appended to the existing path.
' ----------------------------------------------------------------
Public Function filesystemobject_BuildPath(ByVal Path As String _
                                         , ByVal Name As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_BuildPath = m_FSO.BuildPath(Path, Name)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_CopyFile
' Purpose:          Copies one or more files from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter source (String): Required. Character string file specification, which can include wildcard characters, for one or more files to be copied.
' Parameter destination (String): Required. Character string destination where the file or files from source are to be copied. Wildcard characters are not allowed.
' Parameter overwrite (Boolean): Optional. Boolean value that indicates if existing files are to be overwritten. If True, files are overwritten; if False, they are not. The default is True. Note that CopyFile will fail if destination has the read-only attribute set, regardless of the value of overwrite.
' ----------------------------------------------------------------
Public Function filesystemobject_CopyFile(ByVal source As String _
                                        , ByVal destination As String _
                               , Optional ByVal overwrite As Boolean = True)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.CopyFile source, destination, overwrite
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_CopyFolder
' Purpose:          Recursively copies a folder from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter source (String): Required. Character string folder specification, which can include wildcard characters, for one or more folders to be copied.
' Parameter destination (String): Required. Character string destination where the folder and subfolders from source are to be copied. Wildcard characters are not allowed.
' Parameter overwrite (Boolean):    Optional. Boolean value that indicates if existing folders are to be overwritten. If True, files are overwritten; if False, they are not. The default is True.
' ----------------------------------------------------------------
Public Function filesystemobject_CopyFolder(ByVal source As String _
                                          , ByVal destination As String _
                                 , Optional ByVal overwrite As Boolean = True)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.CopyFolder source, destination, overwrite
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_CreateFolder
' Purpose:          Creates a folder.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter foldername (String): Required. String expression that identifies the folder to create.
' ----------------------------------------------------------------
Public Function filesystemobject_CreateFolder(ByVal foldername As String)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.CreateFolder foldername
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_CreateTextFile
' Purpose:          Creates a specified file name and returns a TextStream object that can be used to read from or write to the file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter filename (String): Required. String expression that identifies the file to create.
' Parameter overwrite (Boolean): Optional. Boolean value that indicates if an existing file can be overwritten. The value is True if the file can be overwritten; False if it can't be overwritten. If omitted, existing files can be overwritten.
' Parameter unicode (Boolean): Optional. Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is True if the file is created as a Unicode file; False if it's created as an ASCII file. If omitted, an ASCII file is assumed.
' ----------------------------------------------------------------
Public Function filesystemobject_CreateTextFile(ByVal filename As String _
                                     , Optional ByVal overwrite As Boolean = True _
                                     , Optional ByVal unicode As Boolean = False)
  Dim m_FSO As New my_FileSystemObject_Object

  m_FSO.CreateTextFile filename, overwrite, unicode
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_DeleteFile
' Purpose:          Deletes a specified file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter filespec (String): Required. The name of the file to delete. The filespec can contain wildcard characters in the last path component.
' Parameter force (Boolean):    Optional. Boolean value that is True if files with the read-only attribute set are to be deleted; False (default) if they are not.
' ----------------------------------------------------------------
Public Function filesystemobject_DeleteFile(ByVal filespec As String _
                                 , Optional ByVal force As Boolean = False)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.DeleteFile filespec, force
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_DeleteFolder
' Purpose:          Deletes a specified folder and its contents.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter folderspec (String): Required. The name of the folder to delete. The folderspec can contain wildcard characters in the last path component.
' Parameter force (Boolean): Optional. Boolean value that is True if folders with the read-only attribute set are to be deleted; False (default) if they are not.
' ----------------------------------------------------------------
Public Function filesystemobject_DeleteFolder(ByVal folderspec As String _
                                 , Optional ByVal force As Boolean = False)
  Dim m_FSO As New my_FileSystemObject_Object

  m_FSO.DeleteFolder folderspec, force
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_DriveExists
' Purpose:          Returns True if the specified drive exists; False if it does not.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Boolean
' Parameter drivespec (String): Required. A drive letter or a path specification for the root of the drive.
' ----------------------------------------------------------------
Public Function filesystemobject_DriveExists(ByVal drivespec As String) As Boolean
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_DriveExists = m_FSO.DriveExists(drivespec)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_FileExists
' Purpose:          Returns True if a specified file exists; False if it does not.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Boolean
' Parameter filespec (String): Required. The name of the file whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the file isn't expected to exist in the current folder.
' ----------------------------------------------------------------
Public Function filesystemobject_FileExists(ByVal filespec As String) As Boolean
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_FileExists = m_FSO.FileExists(filespec)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_FolderExists
' Purpose:          Returns True if a specified folder exists; False if it does not.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Boolean
' Parameter folderspec (String): Required. The name of the folder whose existence is to be determined. A complete path specification (either absolute or relative) must be provided if the folder isn't expected to exist in the current folder.
' ----------------------------------------------------------------
Public Function filesystemobject_FolderExists(ByVal folderspec As String) As Boolean
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_FolderExists = m_FSO.FolderExists(folderspec)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetAbsolutePathName
' Purpose:          Returns a complete and unambiguous path from a provided path specification.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter pathspec (String): Required. Path specification to change to a complete and unambiguous path.
' ----------------------------------------------------------------
Public Function filesystemobject_GetAbsolutePathName(ByVal pathspec As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetAbsolutePathName = m_FSO.GetAbsolutePathName(pathspec)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetBaseName
' Purpose:          Returns a string containing the base name of the last component, less any file extension, in a path.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter Path (String): Required. The path specification for the component whose base name is to be returned.
' ----------------------------------------------------------------
Public Function filesystemobject_GetBaseName(ByVal Path As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetBaseName = m_FSO.GetBaseName(Path)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetDrive
' Purpose:          Returns a Drive object corresponding to the drive in a specified path.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' Parameter drivespec (String): Required. The drivespec argument can be a drive letter (c), a drive letter with a colon appended (c:), a drive letter with a colon and path separator appended (c:), or any network share specification (\computer2\share1).
' ----------------------------------------------------------------
Public Function filesystemobject_GetDrive(ByVal drivespec As String) As Object
  Dim m_FSO As New my_FileSystemObject_Object
  
  Set filesystemobject_GetDrive = m_FSO.GetDrive(drivespec)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetDriveName
' Purpose:          Returns a string containing the name of the drive for a specified path.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter Path (String): Required. The path specification for the component whose drive name is to be returned.
' ----------------------------------------------------------------
Public Function filesystemobject_GetDriveName(ByVal Path As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetDriveName = m_FSO.GetDriveName(Path)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetExtensionName
' Purpose:          Returns a string containing the extension name for the last component in a path.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter Path (String): Required. The path specification for the component whose extension name is to be returned.
' ----------------------------------------------------------------
Public Function filesystemobject_GetExtensionName(ByVal Path As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetExtensionName = m_FSO.GetExtensionName(Path)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetFile
' Purpose:          Returns a File object corresponding to the file in a specified path.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' Parameter filespec (String): Required. The filespec is the path (absolute or relative) to a specific file.
' ----------------------------------------------------------------
Public Function filesystemobject_GetFile(ByVal filespec As String) As Object
  Dim m_FSO As New my_FileSystemObject_Object
  
  Set filesystemobject_GetFile = m_FSO.GetFile(filespec)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetFileName
' Purpose:          Returns the last component of a specified path that is not part of the drive specification.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter pathspec (String): Required. The path (absolute or relative) to a specific file.
' ----------------------------------------------------------------
Public Function filesystemobject_GetFileName(ByVal pathspec As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetFileName = m_FSO.GetFileName(pathspec)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetFolder
' Purpose:          Returns a Folder object corresponding to the folder in a specified path.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' Parameter folderspec (String): Required. The folderspec is the path (absolute or relative) to a specific folder.
' ----------------------------------------------------------------
Public Function filesystemobject_GetFolder(ByVal folderspec As String) As Object
  Dim m_FSO As New my_FileSystemObject_Object
    
  Set filesystemobject_GetFolder = m_FSO.GetFolder(folderspec)
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetParentFolderName
' Purpose:          Returns a string containing the name of the parent folder of the last component in a specified path.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter Path (String): Required. The path specification for the component whose parent folder name is to be returned.
' ----------------------------------------------------------------
Public Function filesystemobject_GetParentFolderName(ByVal Path As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetParentFolderName = m_FSO.GetParentFolderName(Path)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetSpecialFolder
' Purpose:          Returns the special folder specified.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter folderspec (String): Required. The name of the special folder to be returned. Can be any of the constants shown in the Settings section.
'                   WindowsFolder   0   The Windows folder contains files installed by the Windows operating system.
'                   SystemFolder    1   The System folder contains libraries, fonts, and device drivers.
'                   TemporaryFolder 2   The Temp folder is used to store temporary files. Its path is found in the TMP environment variable.
' ----------------------------------------------------------------
Public Function filesystemobject_GetSpecialFolder(ByVal folderspec As String) As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetSpecialFolder = m_FSO.GetSpecialFolder(folderspec)
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_GetTempName
' Purpose:          Returns a randomly generated temporary file or folder name that is useful for performing operations that require a temporary file or folder.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Function filesystemobject_GetTempName() As String
  Dim m_FSO As New my_FileSystemObject_Object
  
  filesystemobject_GetTempName = m_FSO.GetTempName
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_MoveFile
' Purpose:          Moves one or more files from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter source (String): Required. The path to the file or files to be moved. The source argument string can contain wildcard characters in the last path component only.
' Parameter destination (String): Required. The path where the file or files are to be moved. The destination argument can't contain wildcard characters.
' ----------------------------------------------------------------
Public Function filesystemobject_MoveFile(ByVal source As String _
                                        , ByVal destination As String)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.MoveFile source, destination
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_MoveFolder
' Purpose:          Moves one or more folders from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter source (String): Required. The path to the folder or folders to be moved. The source argument string can contain wildcard characters in the last path component only.
' Parameter destination (String): Required. The path where the folder or folders are to be moved. The destination argument can't contain wildcard characters.
' ----------------------------------------------------------------
Public Function filesystemobject_MoveFolder(ByVal source As String _
                                          , ByVal destination As String)
  Dim m_FSO As New my_FileSystemObject_Object
  
  m_FSO.MoveFolder source, destination
  
  Set m_FSO = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   filesystemobject_OpenTextFile
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
Public Function filesystemobject_OpenTextFile(ByVal filename As String _
                                   , Optional ByVal mode As my_IOMode = ForAppending _
                                   , Optional ByVal create As Boolean = False _
                                   , Optional ByVal format As my_Tristate = TristateUseDefault) As Object
  Dim m_FSO As New my_FileSystemObject_Object
  
  Set filesystemobject_OpenTextFile = m_FSO.OpenTextFile(filename, mode, create, format)

  Set m_FSO = Nothing
End Function