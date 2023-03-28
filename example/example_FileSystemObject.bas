'
'
' tests/examples for my_FileSystemObject_Object and mdl_FileSystemObject functions
'
'


Private Function test_FileSystemObject_Object()
  ' Verwendung my_FileSystemObject_Object
  
  Dim str_Dir_Acc   As String                                                           ' Variabeln definieren
  Dim str_Dir_New   As String
  Dim str_Dir_Cpy   As String
  Dim str_Fil_New   As String
  Dim str_Fil_Mve   As String
  
  Dim fso           As New my_FileSystemObject_Object                                   ' Objekt instanziieren
  
  str_Dir_Acc = CurrentDb.Name                                                          ' Annahme: Aufruf von Access-Datenbank
  Debug.Print "Aktuelle Access-Datenbank: " & str_Dir_Acc
  
  str_Dir_Acc = fso.GetParentFolderName(str_Dir_Acc)                                    ' Verzeichnis der Datei
  Debug.Print "Verzeichnis Access-Datenbank: " & str_Dir_Acc
  
  str_Dir_New = fso.BuildPath(str_Dir_Acc, "neu_" & format(Now, "yyyymmdd_hhMMss"))     ' Namen für Unterordner erstellen
  Debug.Print "Neues Verzeichnis '" & str_Dir_New & "' vorhanden: " & fso.DriveExists(str_Dir_New)
  
  Debug.Print "Neues Verzeichnis erstellen..."
  fso.CreateFolder str_Dir_New                                                          ' Unterordner erstellen
  
  Debug.Print "Verzeichnis kopieren..."
  str_Dir_Cpy = str_Dir_New & "_Copy"                                                   ' Namen für Verzeichniskopie erstellen
  fso.CopyFolder str_Dir_New, str_Dir_Cpy                                               ' Verzeichniskopie erstellen
  
  Debug.Print "Datei erstellen..."
  str_Fil_New = fso.BuildPath(str_Dir_New, "test.txt")                                  ' Namen für Textdatei erstellen
  fso.CreateTextFile str_Fil_New                                                        ' Textdatei erstellen
  Debug.Print "Datei '" & str_Fil_New & "' vorhanden: " & fso.FileExists(str_Fil_New)
  
  Debug.Print "Datei in Verzeichniskopie verschieben..."
  str_Fil_Mve = str_Dir_Cpy & "/test.txt"
  fso.MoveFile str_Fil_New, str_Fil_Mve                                                 ' Textdatei aus Ursprungsverzeichnis in Verzeichniskopie verschieben
  Debug.Print "Datei '" & str_Fil_Mve & "' vorhanden: " & fso.FileExists(str_Fil_Mve)
  Debug.Print "Datei '" & str_Fil_New & "' vorhanden: " & fso.FileExists(str_Fil_New)
  
  Debug.Print "Datei in Ursprungsverzeichnis kopieren..."
  fso.CopyFile str_Fil_Mve, str_Fil_New                                                 ' Textdatei aus Verzeichniskopie in Ursprungsverzeichnis kopieren
  Debug.Print "Datei '" & str_Fil_Mve & "' vorhanden: " & fso.FileExists(str_Fil_Mve)
  Debug.Print "Datei '" & str_Fil_New & "' vorhanden: " & fso.FileExists(str_Fil_New)
  
  Debug.Print "Verzeichnisse löschen..."
  fso.DeleteFolder str_Dir_New                                                          ' Ursprungsverzeihnis löschen
  fso.DeleteFolder str_Dir_Cpy                                                          ' Verzeichniskopie löschen
  
  Set fso = Nothing
  
End Function


Private Function test_FileSystemObject_Factory()
  ' Verwendung mdl_FileSystemObject
  
  Dim str_Dir_Acc   As String                                                           ' Variabeln definieren
  Dim str_Dir_New   As String
  Dim str_Dir_Cpy   As String
  Dim str_Fil_New   As String
  Dim str_Fil_Mve   As String
  
  str_Dir_Acc = CurrentDb.Name                                                                      ' Annahme: Aufruf von Access-Datenbank
  Debug.Print "Aktuelle Access-Datenbank: " & str_Dir_Acc
  
  str_Dir_Acc = filesystemobject_GetParentFolderName(str_Dir_Acc)                                   ' Verzeichnis der Datei
  Debug.Print "Verzeichnis Access-Datenbank: " & str_Dir_Acc
  
  str_Dir_New = filesystemobject_BuildPath(str_Dir_Acc, "neu_" & format(Now, "yyyymmdd_hhMMss"))    ' Namen für Unterordner erstellen
  Debug.Print "Neues Verzeichnis '" & str_Dir_New & "' vorhanden: " & filesystemobject_DriveExists(str_Dir_New)
  
  Debug.Print "Neues Verzeichnis erstellen..."
  filesystemobject_CreateFolder str_Dir_New                                                         ' Unterordner erstellen
  
  Debug.Print "Verzeichnis kopieren..."
  str_Dir_Cpy = str_Dir_New & "_Copy"                                                               ' Namen für Verzeichniskopie erstellen
  filesystemobject_CopyFolder str_Dir_New, str_Dir_Cpy                                              ' Verzeichniskopie erstellen
  
  Debug.Print "Datei erstellen..."
  str_Fil_New = filesystemobject_BuildPath(str_Dir_New, "test.txt")                                 ' Namen für Textdatei erstellen
  filesystemobject_CreateTextFile str_Fil_New                                                       ' Textdatei erstellen
  Debug.Print "Datei '" & str_Fil_New & "' vorhanden: " & filesystemobject_FileExists(str_Fil_New)
  
  Debug.Print "Datei in Verzeichniskopie verschieben..."
  str_Fil_Mve = str_Dir_Cpy & "/test.txt"
  filesystemobject_MoveFile str_Fil_New, str_Fil_Mve                                                ' Textdatei aus Ursprungsverzeichnis in Verzeichniskopie verschieben
  Debug.Print "Datei '" & str_Fil_Mve & "' vorhanden: " & filesystemobject_FileExists(str_Fil_Mve)
  Debug.Print "Datei '" & str_Fil_New & "' vorhanden: " & filesystemobject_FileExists(str_Fil_New)
  
  Debug.Print "Datei in Ursprungsverzeichnis kopieren..."
  filesystemobject_CopyFile str_Fil_Mve, str_Fil_New                                                ' Textdatei aus Verzeichniskopie in Ursprungsverzeichnis kopieren
  Debug.Print "Datei '" & str_Fil_Mve & "' vorhanden: " & filesystemobject_FileExists(str_Fil_Mve)
  Debug.Print "Datei '" & str_Fil_New & "' vorhanden: " & filesystemobject_FileExists(str_Fil_New)
  
  Debug.Print "Verzeichnisse löschen..."
  filesystemobject_DeleteFolder str_Dir_New                                                         ' Ursprungsverzeihnis löschen
  filesystemobject_DeleteFolder str_Dir_Cpy                                                         ' Verzeichniskopie löschen
  
End Function

  
