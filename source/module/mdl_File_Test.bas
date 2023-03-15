Option Compare Database
Option Explicit
'
'
' tests/examples for mdl_File functions
'
'

Private Function test_01()
  
  Dim datei As New my_File_Object
  
  datei.Name = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
'  datei.Name = "F:\"
  
  Debug.Print "Attributes: " & datei.Attributes
  Debug.Print "DateCreated: " & datei.DateCreated
  Debug.Print "DateLastAccessed: " & datei.DateLastAccessed
  Debug.Print "DateLastModified: " & datei.DateLastModified
  Debug.Print "Drive Letter: " & datei.Drive.DriveLetter
  Debug.Print "Drive FileSystem: " & datei.Drive.FileSystem
  Debug.Print "Name: " & datei.Name
  Debug.Print "ParentFolder Name: " & datei.ParentFolder.Name
  Debug.Print "ParentFolder ShortPath: " & datei.ParentFolder.ShortPath
  Debug.Print "Path: " & datei.Path
  Debug.Print "ShortName: " & datei.ShortName
  Debug.Print "ShortPath: " & datei.ShortPath
  Debug.Print "Size: " & datei.Size
  Debug.Print "Typ: " & datei.Typ
  
 
  Set datei = Nothing
  
End Function

Private Function test_02()
  
  Dim datei As New my_File_Object
  
  Set datei = file_INIT("Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf")
  
'  datei.Name = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
'  datei.Name = "F:\"
  
  Debug.Print "Attributes: " & datei.Attributes
  Debug.Print "DateCreated: " & datei.DateCreated
  Debug.Print "DateLastAccessed: " & datei.DateLastAccessed
  Debug.Print "DateLastModified: " & datei.DateLastModified
  Debug.Print "Drive: " & datei.Drive
  Debug.Print "Name: " & datei.Name
  Debug.Print "ParentFolder: " & datei.ParentFolder
  Debug.Print "Path: " & datei.Path
  Debug.Print "ShortName: " & datei.ShortName
  Debug.Print "ShortPath: " & datei.ShortPath
  Debug.Print "Size: " & datei.Size
  Debug.Print "Typ: " & datei.Typ
  
  Set datei = Nothing
  
End Function

Private Function test_03()
  
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print "Attributes: " & file_Attributes(str_File)
  Debug.Print "DateCreated: " & file_DateCreated(str_File)
  Debug.Print "DateLastAccessed: " & file_DateLastAccessed(str_File)
  Debug.Print "DateLastModified: " & file_DateLastModified(str_File)
  Debug.Print "Drive: " & file_Drive(str_File)
'  Debug.Print "Name: " & file_Name(str_File)
  Debug.Print "ParentFolder: " & file_ParentFolder(str_File)
  Debug.Print "Path: " & file_Path(str_File)
  Debug.Print "ShortName: " & file_ShortName(str_File)
  Debug.Print "ShortPath: " & file_ShortPath(str_File)
  Debug.Print "Size: " & file_Size(str_File)
  Debug.Print "Typ: " & file_Typ(str_File)
  
End Function

Private Function test_File_Info()
  
  Dim str_File As String
  
  str_File = "F:\Datenbank\my_Tool\my_FileSystemObject - Kopie.accdb"
  
  Debug.Print file_ShortName(str_File) & " is a " & file_Typ(str_File)
  
End Function



Private Function test_file_Attributes()
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print file_Attributes(str_File)
  
End Function

Private Function test_file_DateCreated()
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print file_DateCreated(str_File)
  
End Function

Private Function test_file_DateLastAccessed()
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print file_DateLastAccessed(str_File)
  
End Function

Private Function test_file_DateLastModified()
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print file_DateLastModified(str_File)
  
End Function

Private Function test_file_Drive()
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print file_Drive(str_File)
  
End Function

Private Function test_file_ParentFolder()
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print file_ParentFolder(str_File)
  
End Function

Private Function test_file_Path()
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print file_Path(str_File)
  
End Function

Private Function test_file_ShortName()
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print file_ShortName(str_File)
  
End Function

Private Function test_file_ShortPath()
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print file_ShortPath(str_File)
  
End Function

Private Function test_file_Typ()
  Dim str_File As String
  
  str_File = "Z:\Download\AEK\AEK07_Klassen\AEK7_Klassen.pdf"
  
  Debug.Print file_Typ(str_File)
  
End Function