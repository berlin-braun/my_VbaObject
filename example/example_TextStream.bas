'
'
' tests/examples for my_TextStream_Object and mdl_TextStream functions
'
'
Private Function test_TextStream_Object()
  ' Verwendung my_TextStream_Object
  
  Dim str_File    As String
  Dim writer      As New my_TextStream_Object                                           ' Klasse instanziieren
  Dim reader      As New my_TextStream_Object                                           ' Klasse instanziieren
  
  Debug.Print "Schreibe..."
  str_File = CurrentProject.Path & "\text_" & format(Now, "yyyymmdd_hhMMss")            ' Name für Text-File zusammenstellen
  
  writer.CreateTextFile str_File, True, True                                            ' Datei erstellen
  writer.WriteLine "zeile1"                                                             ' Zeile schreiben
  writer.Write_Text "__text__"                                                          ' Zeichen schreiben
  writer.WriteLine "zeile2"                                                             ' Zeile schreiben
  writer.WriteBlankLines 2                                                              ' zwei leere Zeilen schreiben
  writer.WriteLine "zeile3"                                                             ' Zeile schreiben
  writer.WriteLine Now                                                                  ' Zeile schreiben
  writer.Close_Stream
  
  Set writer = Nothing                                                                  ' Objekt schließen
  
  Debug.Print "Lese..."
  
  reader.OpenTextFile str_File, ForReading                                              ' Datei in neues Objekt einlesen
  
  Debug.Print reader.ReadLine                                                           ' eine Zeile auslesen
  Debug.Print reader.ReadLine                                                           ' eine Zeile auslesen
  Debug.Print reader.ReadAll                                                            ' alle weiteren Zeilen auslesen
  
  reader.Close_Stream                                                                   ' Objekt schließen
  
  Set reader = Nothing
  
  file_Delete str_File                                                                  ' Datei löschen
  
End Function

Private Function test_TextStream_Factory()
  ' Verwendung mdl_TextStream
  
  Dim str_File    As String
  
  Debug.Print "Schreibe..."
  str_File = CurrentProject.Path & "\text_" & format(Now, "yyyymmdd_hhMMss")            ' Name für Text-File zusammenstellen
  
  textstream_INIT str_File, ForWriting, True                                            ' Datei erstellen
  textstream_WriteLine str_File, "zeile1"                                               ' Zeile schreiben
  textstream_Write_Text str_File, "__text__"                                            ' Zeichen schreiben
  textstream_WriteLine str_File, "zeile2"                                               ' Zeile schreiben
  textstream_WriteBlankLines str_File, 2                                                ' zwei leere Zeilen schreiben
  textstream_WriteLine str_File, "zeile3"                                               ' Zeile schreiben
  textstream_WriteLine str_File, Now                                                    ' Zeile schreiben
  
  Debug.Print "Lese..."
  Debug.Print textstream_ReadAll(str_File)                                              ' alle Zeilen auslesen
  
  file_Delete str_File                                                                  ' Datei löschen
  
End Function

