Option Compare Database
Option Explicit
'
'
' tests/examples for mdl_Folder functions
'
'

Private Function test_01()
  
  Dim verzeichnis As New my_Folder_Object
  
  verzeichnis.Name = "C:\Users\Thomas Braun\Desktop\Neuer Ordner"
'  verzeichnis.Name = "F:\"
  
  Debug.Print "Attributes: " & verzeichnis.Attributes
  Debug.Print "DateCreated: " & verzeichnis.DateCreated
  Debug.Print "DateLastAccessed: " & verzeichnis.DateLastAccessed
  Debug.Print "DateLastModified: " & verzeichnis.DateLastModified
  Debug.Print "Drive: " & verzeichnis.Drive
  Debug.Print "IsRootFolder: " & verzeichnis.IsRootFolder
  Debug.Print "Name: " & verzeichnis.Name
  Debug.Print "ParentFolder: " & verzeichnis.ParentFolder
  Debug.Print "Path: " & verzeichnis.Path
  Debug.Print "ShortName: " & verzeichnis.ShortName
  Debug.Print "ShortPath: " & verzeichnis.ShortPath
  Debug.Print "Type: " & verzeichnis.Typ
  
  Set verzeichnis = Nothing
  
End Function

Private Function test_02()
  
  Dim verzeichnis As New my_Folder_Object
  
  Set verzeichnis = folder_INIT("C:\Users\Thomas Braun\Desktop\Neuer Ordner")
  
'  verzeichnis.Name = "C:\Users\Thomas Braun\Desktop\Neuer Ordner"
'  verzeichnis.Name = "F:\"
  
  Debug.Print "Attributes: " & verzeichnis.Attributes
  Debug.Print "DateCreated: " & verzeichnis.DateCreated
  Debug.Print "DateLastAccessed: " & verzeichnis.DateLastAccessed
  Debug.Print "DateLastModified: " & verzeichnis.DateLastModified
  Debug.Print "Drive: " & verzeichnis.Drive
  Debug.Print "IsRootFolder: " & verzeichnis.IsRootFolder
  Debug.Print "Name: " & verzeichnis.Name
  Debug.Print "ParentFolder: " & verzeichnis.ParentFolder
  Debug.Print "Path: " & verzeichnis.Path
  Debug.Print "ShortName: " & verzeichnis.ShortName
  Debug.Print "ShortPath: " & verzeichnis.ShortPath
  Debug.Print "Type: " & verzeichnis.Typ
  
  Set verzeichnis = Nothing
  
End Function

Private Function test_03()
  
  Dim verzeichnis As New my_Folder_Object
  
  Set verzeichnis = folder_INIT("C:\tmp\folder_test\02")
  
  verzeichnis.Copy "C:\tmp\folder_test_03\01"
  
  Set verzeichnis = Nothing
  
End Function


Private Function test_04()
  
  Dim verzeichnis As New my_Folder_Object
  
  Set verzeichnis = folder_INIT("C:\tmp\folder_test_02")
  
  verzeichnis.Delete
  
  Set verzeichnis = Nothing
  
End Function

Private Function test_05()
  
  Dim verzeichnis As New my_Folder_Object
  
  Set verzeichnis = folder_INIT("C:\tmp\folder_test\02")
  
  verzeichnis.Move "C:\tmp\folder_test\99"
  
  Set verzeichnis = Nothing
  
End Function



Private Function test_06()
  
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print "Attributes: " & folder_Attributes(str_Folder)
  Debug.Print "DateCreated: " & folder_DateCreated(str_Folder)
  Debug.Print "DateLastAccessed: " & folder_DateLastAccessed(str_Folder)
  Debug.Print "DateLastModified: " & folder_DateLastModified(str_Folder)
  Debug.Print "Drive: " & folder_Drive(str_Folder)
  Debug.Print "IsRootFolder: " & folder_IsRootFolder(str_Folder)
'  Debug.Print "Name: " & folder_Name(str_Folder)
  Debug.Print "ParentFolder: " & folder_ParentFolder(str_Folder)
  Debug.Print "Path: " & folder_Path(str_Folder)
  Debug.Print "ShortName: " & folder_ShortName(str_Folder)
  Debug.Print "ShortPath: " & folder_ShortPath(str_Folder)
  Debug.Print "Type: " & folder_Typ(str_Folder)
  
End Function


Private Function test_folder_Attributes()
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print folder_Attributes(str_Folder)
  
End Function

Private Function test_folder_DateCreated()
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print folder_DateCreated(str_Folder)
  
End Function

Private Function test_folder_DateLastAccessed()
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print folder_DateLastAccessed(str_Folder)
  
End Function

Private Function test_folder_DateLastModified()
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print folder_DateLastModified(str_Folder)
  
End Function

Private Function test_folder_IsRootFolder()
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print folder_IsRootFolder(str_Folder)
  
End Function

Private Function test_folder_ParentFolder()
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print folder_ParentFolder(str_Folder)
  
End Function

Private Function test_folder_Path()
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print folder_Path(str_Folder)
  
End Function

Private Function test_folder_ShortName()
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print folder_ShortName(str_Folder)
  
End Function

Private Function test_folder_ShortPath()
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print folder_ShortPath(str_Folder)
  
End Function

Private Function test_folder_Typ()
  Dim str_Folder As String
  
  str_Folder = "Z:\Download\AEK\AEK07_Klassen"
  
  Debug.Print folder_Typ(str_Folder)
  
End Function