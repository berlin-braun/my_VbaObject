'
'
' tests/examples for mdl_Drive functions
'
'
Private Function test_Drive_Object()
  ' Verwendung my_Drive_Object

  Dim str_Drive   As String
  Dim str_Ret     As String
  
  str_Drive = InputBox("Bitte Laufwerksbuchstaben eingeben." _
                     , "Example")                                               ' Laufwerksbuchstaben abfragen
  
  If Len(str_Drive) > 0 Then                                                    ' Eingabe erfolgt
    
    str_Drive = Left(Trim(str_Drive), 1)                                        ' nur erstes Zeichen
    
    If IsNumeric(str_Drive) = False Then                                        ' keine Zahl
        
      Dim Drive As New my_Drive_Object                                          ' Klasse instanziieren
      
      Set Drive = drive_INIT(str_Drive)                                         ' Laufwerksbuchstaben übergeben
      
      str_Ret = str_Ret & vbCrLf & "AvailableSpace: " & Drive.AvailableSpace    ' Informationen sammeln
      str_Ret = str_Ret & vbCrLf & "DriveLetter: " & Drive.DriveLetter
      str_Ret = str_Ret & vbCrLf & "DriveType: " & Drive.DriveType
      str_Ret = str_Ret & vbCrLf & "FileSystem: " & Drive.FileSystem
      str_Ret = str_Ret & vbCrLf & "FreeSpace: " & Drive.FreeSpace
      str_Ret = str_Ret & vbCrLf & "IsReady: " & Drive.IsReady
      str_Ret = str_Ret & vbCrLf & "Path: " & Drive.Path
      str_Ret = str_Ret & vbCrLf & "RootFolder: " & Drive.RootFolder
      str_Ret = str_Ret & vbCrLf & "SerialNumber: " & Drive.SerialNumber
      str_Ret = str_Ret & vbCrLf & "ShareName: " & Drive.ShareName
      str_Ret = str_Ret & vbCrLf & "TotalSize: " & Drive.TotalSize
      str_Ret = str_Ret & vbCrLf & "VolumeName: " & Drive.VolumeName
      
      Debug.Print str_Ret                                                       ' Informationen ausgeben
      MsgBox str_Ret                                                            ' Information anzeigen
      
      Set Drive = Nothing
      
    End If
    
  End If
  
End Function


Private Function test_Drive_Factory()
  ' Verwendung mdl_Drive

  Dim str_Drive   As String
  Dim str_Ret     As String
  
  str_Drive = InputBox("Bitte Laufwerksbuchstaben eingeben." _
                     , "Example")                                                         ' Laufwerksbuchstaben abfragen
  
  If Len(str_Drive) > 0 Then                                                              ' Eingabe erfolgt
    
    str_Drive = Left(Trim(str_Drive), 1)                                                  ' nur erstes Zeichen
    
    If IsNumeric(str_Drive) = False Then                                                  ' keine Zahl
        
      str_Ret = str_Ret & vbCrLf & "AvailableSpace: " & drive_AvailableSpace(str_Drive)   ' Informationen sammeln
      str_Ret = str_Ret & vbCrLf & "DriveLetter: " & str_Drive
      str_Ret = str_Ret & vbCrLf & "DriveType: " & drive_DriveType(str_Drive)
      str_Ret = str_Ret & vbCrLf & "FileSystem: " & drive_FileSystem(str_Drive)
      str_Ret = str_Ret & vbCrLf & "FreeSpace: " & drive_FreeSpace(str_Drive)
      str_Ret = str_Ret & vbCrLf & "IsReady: " & drive_IsReady(str_Drive)
      str_Ret = str_Ret & vbCrLf & "Path: " & drive_Path(str_Drive)
      str_Ret = str_Ret & vbCrLf & "RootFolder: " & drive_RootFolder(str_Drive)
      str_Ret = str_Ret & vbCrLf & "SerialNumber: " & drive_SerialNumber(str_Drive)
      str_Ret = str_Ret & vbCrLf & "ShareName: " & drive_ShareName(str_Drive)
      str_Ret = str_Ret & vbCrLf & "TotalSize: " & drive_TotalSize(str_Drive)
      str_Ret = str_Ret & vbCrLf & "VolumeName: " & drive_VolumeName(str_Drive)
      
      Debug.Print str_Ret                                                                 ' Informationen ausgeben
      MsgBox str_Ret                                                                      ' Information anzeigen
      
    End If
    
  End If
  
End Function


