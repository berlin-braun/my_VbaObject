Option Compare Database
Option Explicit
'
'
' Factory: static properties and methods for class my_Drive_Object
'
'


Public Function drive_INIT(ByVal drive_Letter As String) As my_Drive_Object
  Dim m_Drive     As New my_Drive_Object
  
  m_Drive.DriveLetter = drive_Letter
  Set drive_INIT = m_Drive
  
  Set m_Drive = Nothing
End Function

' ----------------------------------------------------------------
' Procedure Name:   drive_AvailableSpace
' Purpose:          Returns the amount of space available to a user on the specified drive or network share.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' Parameter str_Drive_Letter (String): Letter of a specific drive
' ----------------------------------------------------------------
Public Function drive_AvailableSpace(ByVal str_Drive_Letter As String) As Long
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_AvailableSpace = m_Drive.AvailableSpace
  
  Set m_Drive = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   drive_DriveType
' Purpose:          Returns a value indicating the type of a specified drive.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      my_DriveType
' Parameter str_Drive_Letter (String): Letter of the specific drive
' ----------------------------------------------------------------
Public Function drive_DriveType(ByVal str_Drive_Letter As String) As my_DriveType
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_DriveType = m_Drive.DriveType
  
  Set m_Drive = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   drive_FileSystem
' Purpose:          Returns the type of file system in use for the specified drive.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter str_Drive_Letter (String): Letter of the specific drive
' ----------------------------------------------------------------
Public Function drive_FileSystem(ByVal str_Drive_Letter As String) As String
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_FileSystem = m_Drive.FileSystem
  
  Set m_Drive = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   drive_FreeSpace
' Purpose:          Returns the amount of free space available to a user on the specified drive or network share. Read-only.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' Parameter str_Drive_Letter (String): Letter of the specific drive
' ----------------------------------------------------------------
Public Function drive_FreeSpace(ByVal str_Drive_Letter As String) As Long
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_FreeSpace = m_Drive.FreeSpace
  
  Set m_Drive = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   drive_IsReady
' Purpose:          Returns True if the specified drive is ready; False if it is not.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Boolean
' Parameter str_Drive_Letter (String): Letter of the specific drive
' ----------------------------------------------------------------
Public Function drive_IsReady(ByVal str_Drive_Letter As String) As Boolean
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_IsReady = m_Drive.IsReady
  
  Set m_Drive = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   drive_Path
' Purpose:
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter str_Drive_Letter (String): Letter of the specific drive
' ----------------------------------------------------------------
Public Function drive_Path(ByVal str_Drive_Letter As String) As String
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_Path = m_Drive.Path
  
  Set m_Drive = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   drive_RootFolder
' Purpose:          Returns a Folder object representing the root folder of a specified drive. Read-only.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' Parameter str_Drive_Letter (String): Letter of the specific drive
' ----------------------------------------------------------------
Public Function drive_RootFolder(ByVal str_Drive_Letter As String) As Object
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  Set drive_RootFolder = m_Drive.RootFolder
  
  Set m_Drive = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   drive_SerialNumber
' Purpose:          Returns the decimal serial number used to uniquely identify a disk volume.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' Parameter str_Drive_Letter (String): Letter of the specific drive
' ----------------------------------------------------------------
Public Function drive_SerialNumber(ByVal str_Drive_Letter As String) As Long
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_SerialNumber = m_Drive.SerialNumber
  
  Set m_Drive = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   drive_ShareName
' Purpose:          Returns the network share name for a specified drive.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter str_Drive_Letter (String): Letter of the specific drive
' ----------------------------------------------------------------
Public Function drive_ShareName(ByVal str_Drive_Letter As String) As String
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_ShareName = m_Drive.ShareName
  
  Set m_Drive = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   drive_TotalSize
' Purpose:          Returns the total space, in bytes, of a drive or network share.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' Parameter str_Drive_Letter (String): Letter of the specific drive
' ----------------------------------------------------------------
Public Function drive_TotalSize(ByVal str_Drive_Letter As String) As Long
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_TotalSize = m_Drive.TotalSize
  
  Set m_Drive = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   drive_VolumeName
' Purpose:          Returns the volume name of the specified drive. Read/write.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter str_Drive_Letter (String): Letter of the specific drive
' ----------------------------------------------------------------
Public Function drive_VolumeName(ByVal str_Drive_Letter As String) As String
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_VolumeName = m_Drive.VolumeName
  
  Set m_Drive = Nothing
End Function