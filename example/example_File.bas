'
'
' tests/examples for my_File_Object and mdl_File functions
'
'
Private Function test_File_Object()
  ' Verwendung my_File_Object
  
  Dim str_Ret     As String
  Dim datei       As New my_File_Object                                               ' Klasse instanziieren
  
  datei.Name = CurrentDb.Name                                                         ' Annahme: Aufruf aus Access-Daternbank
  
  str_Ret = str_Ret & vbCrLf & "Attributes: " & datei.Attributes                      ' Informationen sammeln
  str_Ret = str_Ret & vbCrLf & "DateCreated: " & datei.DateCreated
  str_Ret = str_Ret & vbCrLf & "DateLastAccessed: " & datei.DateLastAccessed
  str_Ret = str_Ret & vbCrLf & "DateLastModified: " & datei.DateLastModified
  str_Ret = str_Ret & vbCrLf & "Drive Letter: " & datei.Drive.DriveLetter
  str_Ret = str_Ret & vbCrLf & "Drive FileSystem: " & datei.Drive.FileSystem
  str_Ret = str_Ret & vbCrLf & "Name: " & datei.Name
  str_Ret = str_Ret & vbCrLf & "ParentFolder Name: " & datei.ParentFolder.Name
  str_Ret = str_Ret & vbCrLf & "ParentFolder ShortPath: " & datei.ParentFolder.ShortPath
  str_Ret = str_Ret & vbCrLf & "Path: " & datei.Path
  str_Ret = str_Ret & vbCrLf & "ShortName: " & datei.ShortName
  str_Ret = str_Ret & vbCrLf & "ShortPath: " & datei.ShortPath
  str_Ret = str_Ret & vbCrLf & "Size: " & datei.Size
  str_Ret = str_Ret & vbCrLf & "Typ: " & datei.Typ
  
  Debug.Print str_Ret                                                                 ' Informationen ausgeben
  MsgBox str_Ret                                                                      ' Informationen anzeigen
  
  Set datei = Nothing
  
End Function

Private Function test_File_Factory()
  ' Verwendung mdl_File
  
  Dim str_Ret     As String
  Dim str_File    As String
  
  str_File = CurrentDb.Name                                                           ' Annahme: Aufruf aus Access-Daternbank
  
  str_Ret = str_Ret & vbCrLf & "Attributes: " & file_Attributes(str_File)             ' Informationen mit den Factory-Methoden sammeln
  str_Ret = str_Ret & vbCrLf & "DateCreated: " & file_DateCreated(str_File)
  str_Ret = str_Ret & vbCrLf & "DateLastAccessed: " & file_DateLastAccessed(str_File)
  str_Ret = str_Ret & vbCrLf & "DateLastModified: " & file_DateLastModified(str_File)
  str_Ret = str_Ret & vbCrLf & "Drive: " & file_Drive(str_File)
  str_Ret = str_Ret & vbCrLf & "ParentFolder: " & file_ParentFolder(str_File)
  str_Ret = str_Ret & vbCrLf & "Path: " & file_Path(str_File)
  str_Ret = str_Ret & vbCrLf & "ShortName: " & file_ShortName(str_File)
  str_Ret = str_Ret & vbCrLf & "ShortPath: " & file_ShortPath(str_File)
  str_Ret = str_Ret & vbCrLf & "Size: " & file_Size(str_File)
  str_Ret = str_Ret & vbCrLf & "Typ: " & file_Typ(str_File)
  
  Debug.Print str_Ret                                                                 ' Informationen ausgeben
  MsgBox str_Ret                                                                      ' Informationen anzeigen
  
End Function

