Option Compare Database
Option Explicit
'
'
' tests/examples for mdl_FileSystemObject functions
'
'



Private Function test_Drives()
  Dim lng_Drives As Long
  
  lng_Drives = 10                               ' harddrive + cd-rom
  
  Debug.Print Drives.Count
  Debug.Assert Drives.Count = lng_Drives
  
End Function


Private Function test_BuildPath()
  Dim Path    As String
  Dim Name    As String
  Dim result  As String
  
  Path = "C:\test"
  Name = "new.txt"
  result = filesystemobject_BuildPath(Path, Name)
  
  Debug.Print result
  Debug.Assert result = Path & "\" & Name          ' soll "C:\test\new.txt"

End Function

Private Function test_CopyFile()
  Dim source        As String
  Dim destination   As String
  Dim overwrite     As Boolean
  
  source = "F:\Datenbank\my_Tool\test\test.txt"
  destination = "F:\Datenbank\my_Tool\test\test_" & format(Now, "yyyyMMdd_hhnnss") & ".txt"
  overwrite = True
  
  filesystemobject_CopyFile source, destination, overwrite
  
  Debug.Print destination
  Debug.Assert Dir(destination) <> ""
  
End Function

Private Function test_CopyFolder()
  Dim source        As String
  Dim destination   As String
  Dim overwrite     As Boolean
  
  source = "F:\Datenbank\my_Tool\test"
  destination = "F:\Datenbank\my_Tool\test_" & format(Now, "yyyyMMdd_hhnnss")
  overwrite = True
  
  filesystemobject_CopyFolder source, destination, overwrite
  
  Debug.Print destination
  Debug.Assert Dir(destination & "\*.*") <> ""
  
End Function

Private Function test_CreateFolder()
  Dim foldername    As String
  
  foldername = "F:\Datenbank\my_Tool\test_create_" & format(Now, "yyyyMMdd_hhnnss")
  filesystemobject_CreateFolder foldername
  
  Debug.Print foldername
  
End Function

Private Function test_CreateTextFile()
  Dim filename      As String
  Dim overwrite     As Boolean
  Dim unicode       As Boolean
  
  overwrite = True
  unicode = False
  filename = "F:\Datenbank\my_Tool\test\test_create_" & format(Now, "yyyyMMdd_hhnnss") & ".txt"

  filesystemobject_CreateTextFile filename, overwrite, unicode
  
  Debug.Print filename
  Debug.Assert Dir(filename) <> ""

End Function

Private Function test_DeleteFile()
  Dim filespec      As String
  Dim force         As Boolean
  
  force = False
  filespec = "F:\Datenbank\my_Tool\test\test_create_" & format(Now, "yyyyMMdd_hhnnss") & ".txt"

  filesystemobject_CreateTextFile filespec
  
  Debug.Print filespec
  Debug.Assert Dir(filespec) <> ""
 
  filesystemobject_DeleteFile filespec, force
  
  Debug.Print filespec
  Debug.Assert Dir(filespec) = ""

End Function

Private Function test_DeleteFolder()
  Dim folderspec    As String
  Dim force         As Boolean

  force = False
  
  folderspec = "F:\Datenbank\my_Tool\test\test_delete_" & format(Now, "yyyyMMdd_hhnnss")
  filesystemobject_CreateFolder folderspec
  
  Debug.Print folderspec
  
  filesystemobject_DeleteFolder folderspec, force
  
  Debug.Print folderspec

End Function

Private Function test_DriveExists()
  Dim drivespec As String
  
  drivespec = "C"
  
  Debug.Print filesystemobject_DriveExists(drivespec)

End Function

Private Function test_FileExists()
  Dim filespec As String
  
  filespec = "F:\Datenbank\my_Tool\test\test_exists_" & format(Now, "yyyyMMdd_hhnnss") & ".txt"
  
  Debug.Print filesystemobject_FileExists(filespec)
  
  filesystemobject_CreateTextFile filespec, True
  Debug.Print filesystemobject_FileExists(filespec)
  
  filesystemobject_DeleteFile filespec
  Debug.Print filesystemobject_FileExists(filespec)
  
End Function

Private Function test_FolderExists()
  Dim folderspec As String
  
  folderspec = "F:\Datenbank\my_Tool\test\test_exists_" & format(Now, "yyyyMMdd_hhnnss")
  
  Debug.Print filesystemobject_FolderExists(folderspec)
  
  filesystemobject_CreateFolder folderspec
  Debug.Print filesystemobject_FolderExists(folderspec)
  
  filesystemobject_DeleteFolder folderspec
  Debug.Print filesystemobject_FolderExists(folderspec)
  
End Function

Private Function test_GetAbsolutePathName()
  Dim pathspec As String
  
  pathspec = "F:\..\my_Tool\test"
  
  Debug.Print filesystemobject_GetAbsolutePathName(pathspec)
  

End Function

Private Function test_GetBaseName()
  Dim Path As String
  
  Path = CurrentProject.Path
  
  Debug.Print filesystemobject_GetBaseName(Path)

End Function

Private Function test_GetDrive()
  Dim drivespec As String
  
  drivespec = "F"
  
  Debug.Print filesystemobject_GetDrive(drivespec).DriveType
  Debug.Print filesystemobject_GetDrive(drivespec).FileSystem

End Function

Private Function test_GetDriveName()
  Dim Path As String
  
  Path = "F:\Datenbank\my_Tool\test"
  
  Debug.Print filesystemobject_GetDriveName(Path)
  
End Function

Private Function test_GetExtensionName()
  Dim Path As String
  
  Path = CurrentProject.Name
  
  Debug.Print filesystemobject_GetExtensionName(Path)
  
End Function

Private Function test_GetFile()
  Dim filespec As String
  
  filespec = CurrentDb.Name
  
  Debug.Print filesystemobject_GetFile(filespec).DateCreated
  Debug.Print filesystemobject_GetFile(filespec).Size
  
End Function

Private Function test_GetFileName()
  Dim pathspec As String
  
  pathspec = "F:\Datenbank\my_Tool\my_FileSystemObject.accdb"
  
  Debug.Print filesystemobject_GetFileName(pathspec)

End Function

Private Function test_GetFolder()
  Dim folderspec As String
    
  folderspec = "F:\Datenbank\my_Tool"
  
  Debug.Print filesystemobject_GetFolder(folderspec).Attributes
  Debug.Print filesystemobject_GetFolder(folderspec).DateLastAccessed
  Debug.Print filesystemobject_GetFolder(folderspec).Drive
  Debug.Print filesystemobject_GetFolder(folderspec).Path
  
End Function

Private Function test_GetParentFolderName()
  Dim Path As String
  
  Path = "F:\Datenbank\my_Tool\"
  
  Debug.Print filesystemobject_GetParentFolderName(Path)
  
End Function

Private Function test_GetSpecialFolder()
  Dim folderspec As String
  
  Debug.Print filesystemobject_GetSpecialFolder("0")   ' WindowsFolder
  Debug.Print filesystemobject_GetSpecialFolder("1")   ' SystemFolder
  Debug.Print filesystemobject_GetSpecialFolder("2")   ' TemporaryFolder
  
End Function

Private Function test_GetTempName() As String
    
  Debug.Print filesystemobject_GetTempName
  
End Function

Private Function test_MoveFile()
  Dim source            As String
  Dim destination       As String
  
  source = "F:\Datenbank\my_Tool\test\test.txt"
  destination = "F:\Datenbank\my_Tool\test\01\test-2.txt"
  
  filesystemobject_MoveFile source, destination
  
  
End Function

Private Function test_MoveFolder()
  Dim source            As String
  Dim destination       As String
  
  source = "F:\Datenbank\my_Tool\test\01"
  destination = "F:\Datenbank\my_Tool\test\03\01"
  
  filesystemobject_MoveFolder source, destination
  
End Function

Private Function test_OpenTextFile()
  Dim filename          As String
  Dim mode              As IOMode
  Dim create            As Boolean
  Dim format            As Tristate
  
  Dim file As Object
  Dim text As String
  
  text = "hello" & VBA.format$(Now, "yyyyMMdd_hhnnss") & vbCrLf
  
  filename = "F:\Datenbank\my_Tool\test\test.txt"
  mode = ForAppending
  create = False
  format = TristateUseDefault
  
  Set file = filesystemobject_OpenTextFile(filename, mode, create, format)
  file.Write text
  file.Close
  
  Set file = Nothing

End Function