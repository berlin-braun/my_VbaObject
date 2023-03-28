'
'
' tests/examples for my_Folder_Object and mdl_Folder functions
'
'
Private Function test_Folder_Object()
  ' Verwendung my_Folder_Object
  
  Dim str_Ret       As String
  Dim str_Folder    As String
  
  Dim folder        As New my_Folder_Object
  
  str_Folder = CurrentProject.Path                                                    ' Annahme: Aufruf aus Access-Daternbank
  
  Set folder = folder_INIT(str_Folder)                                                ' Klasse instanziieren
  
  str_Ret = str_Ret & vbCrLf & "Attributes: " & folder.Attributes                     ' Informationen mit den Klassen-Methoden sammeln
  str_Ret = str_Ret & vbCrLf & "DateCreated: " & folder.DateCreated
  str_Ret = str_Ret & vbCrLf & "DateLastAccessed: " & folder.DateLastAccessed
  str_Ret = str_Ret & vbCrLf & "DateLastModified: " & folder.DateLastModified
  str_Ret = str_Ret & vbCrLf & "Drive: " & folder.Drive
  str_Ret = str_Ret & vbCrLf & "IsRootFolder: " & folder.IsRootFolder
  str_Ret = str_Ret & vbCrLf & "ParentFolder: " & folder.ParentFolder
  str_Ret = str_Ret & vbCrLf & "Path: " & folder.Path
  str_Ret = str_Ret & vbCrLf & "ShortName: " & folder.ShortName
  str_Ret = str_Ret & vbCrLf & "ShortPath: " & folder.ShortPath
  str_Ret = str_Ret & vbCrLf & "Type: " & folder.Typ
  
  Set folder = Nothing
  
  Debug.Print str_Ret                                                                 ' Informationen ausgeben
  MsgBox str_Ret                                                                      ' Informationen anzeigen
  
End Function

Private Function test_Folder_Factory()
  ' Verwendung mdl_Folder
  
  Dim str_Ret       As String
  Dim str_Folder    As String
  
  str_Folder = CurrentProject.Path                                                    ' Annahme: Aufruf aus Access-Daternbank
  
  str_Ret = str_Ret & vbCrLf & "Attributes: " & folder_Attributes(str_Folder)         ' Informationen mit den Factory-Methoden sammeln
  str_Ret = str_Ret & vbCrLf & "DateCreated: " & folder_DateCreated(str_Folder)
  str_Ret = str_Ret & vbCrLf & "DateLastAccessed: " & folder_DateLastAccessed(str_Folder)
  str_Ret = str_Ret & vbCrLf & "DateLastModified: " & folder_DateLastModified(str_Folder)
  str_Ret = str_Ret & vbCrLf & "Drive: " & folder_Drive(str_Folder)
  str_Ret = str_Ret & vbCrLf & "IsRootFolder: " & folder_IsRootFolder(str_Folder)
  str_Ret = str_Ret & vbCrLf & "ParentFolder: " & folder_ParentFolder(str_Folder)
  str_Ret = str_Ret & vbCrLf & "Path: " & folder_Path(str_Folder)
  str_Ret = str_Ret & vbCrLf & "ShortName: " & folder_ShortName(str_Folder)
  str_Ret = str_Ret & vbCrLf & "ShortPath: " & folder_ShortPath(str_Folder)
  str_Ret = str_Ret & vbCrLf & "Type: " & folder_Typ(str_Folder)
  
  Debug.Print str_Ret                                                                 ' Informationen ausgeben
  MsgBox str_Ret                                                                      ' Informationen anzeigen
  
End Function

