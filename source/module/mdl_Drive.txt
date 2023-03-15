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


Public Function drive_AvailableSpace(ByVal str_Drive_Letter As String) As Long
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_AvailableSpace = m_Drive.AvailableSpace
  
  Set m_Drive = Nothing
End Function

Public Function drive_DriveType(ByVal str_Drive_Letter As String) As DriveTypeConst
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_DriveType = m_Drive.DriveType
  
  Set m_Drive = Nothing
End Function

Public Function drive_FileSystem(ByVal str_Drive_Letter As String) As String
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_FileSystem = m_Drive.FileSystem
  
  Set m_Drive = Nothing
End Function

Public Function drive_FreeSpace(ByVal str_Drive_Letter As String) As Long
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_FreeSpace = m_Drive.FreeSpace
  
  Set m_Drive = Nothing
End Function

Public Function drive_IsReady(ByVal str_Drive_Letter As String) As Boolean
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_IsReady = m_Drive.IsReady
  
  Set m_Drive = Nothing
End Function

Public Function drive_Path(ByVal str_Drive_Letter As String) As String
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_Path = m_Drive.Path
  
  Set m_Drive = Nothing
End Function

Public Function drive_RootFolder(ByVal str_Drive_Letter As String) As Object
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  Set drive_RootFolder = m_Drive.RootFolder
  
  Set m_Drive = Nothing
End Function

Public Function drive_SerialNumber(ByVal str_Drive_Letter As String) As Long
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_SerialNumber = m_Drive.SerialNumber
  
  Set m_Drive = Nothing
End Function

Public Function drive_ShareName(ByVal str_Drive_Letter As String) As String
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_ShareName = m_Drive.ShareName
  
  Set m_Drive = Nothing
End Function

Public Function drive_TotalSize(ByVal str_Drive_Letter As String) As Long
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_TotalSize = m_Drive.TotalSize
  
  Set m_Drive = Nothing
End Function

Public Function drive_VolumeName(ByVal str_Drive_Letter As String) As String
  Dim m_Drive As New my_Drive_Object
  
  Set m_Drive = drive_INIT(str_Drive_Letter)
  
  drive_VolumeName = m_Drive.VolumeName
  
  Set m_Drive = Nothing
End Function