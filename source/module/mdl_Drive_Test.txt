Option Compare Database
Option Explicit
'
'
' tests/examples for mdl_Drive functions
'
'

Private Function test_01()
  
  Dim Drive As New my_Drive_Object
  
  Drive.DriveLetter = "H"
  
  Debug.Print Drive.AvailableSpace
  Debug.Print Drive.DriveLetter
  Debug.Print Drive.DriveType
  Debug.Print Drive.FileSystem
  Debug.Print Drive.FreeSpace
  Debug.Print Drive.IsReady
  Debug.Print Drive.Path
  Debug.Print Drive.RootFolder
  Debug.Print Drive.SerialNumber
  Debug.Print Drive.ShareName
  Debug.Print Drive.TotalSize
  Debug.Print Drive.VolumeName
  
End Function
  
Private Function test_02()
  
  Dim Drive As New my_Drive_Object
  
  Set Drive = drive_INIT("F")
  
  Debug.Print "AvailableSpace: " & Drive.AvailableSpace
  Debug.Print "DriveLetter: " & Drive.DriveLetter
  Debug.Print "DriveType: " & Drive.DriveType
  Debug.Print "FileSystem: " & Drive.FileSystem
  Debug.Print "FreeSpace: " & Drive.FreeSpace
  Debug.Print "IsReady: " & Drive.IsReady
  Debug.Print "Path: " & Drive.Path
  Debug.Print "RootFolder: " & Drive.RootFolder
  Debug.Print "SerialNumber: " & Drive.SerialNumber
  Debug.Print "ShareName: " & Drive.ShareName
  Debug.Print "TotalSize: " & Drive.TotalSize
  Debug.Print "VolumeName: " & Drive.VolumeName
  
  Set Drive = Nothing
  
End Function

Private Function test_03()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print "AvailableSpace: " & drive_AvailableSpace(str_Drive_Letter)
  Debug.Print "DriveLetter: " & str_Drive_Letter
  Debug.Print "DriveType: " & drive_DriveType(str_Drive_Letter)
  Debug.Print "FileSystem: " & drive_FileSystem(str_Drive_Letter)
  Debug.Print "FreeSpace: " & drive_FreeSpace(str_Drive_Letter)
  Debug.Print "IsReady: " & drive_IsReady(str_Drive_Letter)
  Debug.Print "Path: " & drive_Path(str_Drive_Letter)
  Debug.Print "RootFolder: " & drive_RootFolder(str_Drive_Letter)
  Debug.Print "SerialNumber: " & drive_SerialNumber(str_Drive_Letter)
  Debug.Print "ShareName: " & drive_ShareName(str_Drive_Letter)
  Debug.Print "TotalSize: " & drive_TotalSize(str_Drive_Letter)
  Debug.Print "VolumeName: " & drive_VolumeName(str_Drive_Letter)
  
End Function



Private Function test_drive_AvailableSpace()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_AvailableSpace(str_Drive_Letter)
  
End Function

Private Function test_drive_DriveType()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_DriveType(str_Drive_Letter)

End Function

Private Function test_drive_FileSystem()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_FileSystem(str_Drive_Letter)

End Function

Private Function test_drive_FreeSpace()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_FreeSpace(str_Drive_Letter)

End Function

Private Function test_drive_IsReady()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_IsReady(str_Drive_Letter)

End Function

Private Function test_drive_Path()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_Path(str_Drive_Letter)

End Function

Private Function test_drive_RootFolder()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_RootFolder(str_Drive_Letter)

End Function

Private Function test_drive_SerialNumber()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_SerialNumber(str_Drive_Letter)

End Function

Private Function test_drive_ShareName()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_ShareName(str_Drive_Letter)

End Function

Private Function test_drive_TotalSize()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_TotalSize(str_Drive_Letter)

End Function

Private Function test_drive_VolumeName()
  Dim str_Drive_Letter As String
  
  str_Drive_Letter = "C"
  
  Debug.Print drive_VolumeName(str_Drive_Letter)

End Function