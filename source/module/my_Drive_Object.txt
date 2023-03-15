Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum my_DriveType
  Unknown = 0
  Removable = 1
  Fixed = 2
  Remote = 3
  CD_ROM = 4
  RAM_Disk = 5
End Enum

Private my_fso As Scripting.FileSystemObject
Private my_drv As Object
'


Private Sub Class_Terminate()
  
  Set my_fso = Nothing
  Set my_drv = Nothing

End Sub


' Properties - Start

Public Property Get AvailableSpace() As Long
  AvailableSpace = FormatNumber(my_drv.AvailableSpace / 1024, 0)
End Property

Public Property Get DriveLetter() As String
  DriveLetter = my_drv.DriveLetter
End Property

Public Property Let DriveLetter(ByVal str_Drive As String)
  
  Set my_fso = Nothing
  Set my_fso = New Scripting.FileSystemObject
  
  Set my_drv = my_fso.GetDrive(UCase(str_Drive))
  
End Property

Public Property Get DriveType() As DriveTypeConst
  DriveType = my_drv.DriveType
End Property

Public Property Get FileSystem() As String
  FileSystem = my_drv.FileSystem
End Property

Public Property Get FreeSpace() As Long
  FreeSpace = FormatNumber(my_drv.FreeSpace / 1024, 0)
End Property

Public Property Get IsReady() As Boolean
  IsReady = my_drv.IsReady
End Property

Public Property Get Path() As String
  Path = my_drv.Path
End Property

Public Property Get RootFolder() As Object
  Set RootFolder = my_drv.RootFolder
End Property

Public Property Get SerialNumber() As Long
  SerialNumber = my_drv.SerialNumber
End Property

Public Property Get ShareName() As String
  ShareName = my_drv.ShareName
End Property

Public Property Get TotalSize() As Long
  TotalSize = FormatNumber(my_drv.TotalSize / 1024, 0)
End Property

Public Property Get VolumeName() As String
  VolumeName = my_drv.VolumeName
End Property

' Properties - End