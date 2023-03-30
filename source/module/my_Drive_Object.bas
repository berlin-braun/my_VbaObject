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

Private my_fso As Object ' Scripting.FileSystemObject
Private my_drv As Object
'


Private Sub Class_Terminate()
  
  Set my_fso = Nothing
  Set my_drv = Nothing

End Sub


' Properties - Start


' ----------------------------------------------------------------
' Procedure Name:   AvailableSpace
' Purpose:          Returns the amount of space available to a user on the specified drive or network share.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' ----------------------------------------------------------------
Public Property Get AvailableSpace() As Long
  AvailableSpace = FormatNumber(my_drv.AvailableSpace / 1024, 0)
End Property



' ----------------------------------------------------------------
' Procedure Name:   DriveLetter
' Purpose:          Returns the drive letter of a physical local drive or a network share. Read-only.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get DriveLetter() As String
  DriveLetter = my_drv.DriveLetter
End Property


' ----------------------------------------------------------------
' Procedure Name:   DriveLetter
' Purpose:          Sets and initiates the drive letter of a physical local drive or a network share
' Procedure Kind:   Eigenschaft (Let)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter str_Drive (String):
' ----------------------------------------------------------------
Public Property Let DriveLetter(ByVal str_Drive As String)
  
  Set my_fso = Nothing
  Set my_fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
  
  Set my_drv = my_fso.GetDrive(UCase(str_Drive))
  
End Property


' ----------------------------------------------------------------
' Procedure Name:   DriveType
' Purpose:          Returns a value indicating the type of a specified drive.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      my_DriveType
' ----------------------------------------------------------------
Public Property Get DriveType() As my_DriveType ' DriveTypeConst
  DriveType = my_drv.DriveType
End Property


' ----------------------------------------------------------------
' Procedure Name:   FileSystem
' Purpose:          Returns the type of file system in use for the specified drive.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get FileSystem() As String
  FileSystem = my_drv.FileSystem
End Property


' ----------------------------------------------------------------
' Procedure Name:   FreeSpace
' Purpose:          Returns the amount of free space available to a user on the specified drive or network share. Read-only.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' ----------------------------------------------------------------
Public Property Get FreeSpace() As Long
  FreeSpace = FormatNumber(my_drv.FreeSpace / 1024, 0)
End Property


' ----------------------------------------------------------------
' Procedure Name:   IsReady
' Purpose:          Returns True if the specified drive is ready; False if it is not.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Boolean
' ----------------------------------------------------------------
Public Property Get IsReady() As Boolean
  IsReady = my_drv.IsReady
End Property


' ----------------------------------------------------------------
' Procedure Name:   Path
' Purpose:
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get Path() As String
  Path = my_drv.Path
End Property


' ----------------------------------------------------------------
' Procedure Name:   RootFolder
' Purpose:          Returns a Folder object representing the root folder of a specified drive. Read-only.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' ----------------------------------------------------------------
Public Property Get RootFolder() As Object
  Set RootFolder = my_drv.RootFolder
End Property


' ----------------------------------------------------------------
' Procedure Name:   SerialNumber
' Purpose:          Returns the decimal serial number used to uniquely identify a disk volume.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' ----------------------------------------------------------------
Public Property Get SerialNumber() As Long
  SerialNumber = my_drv.SerialNumber
End Property


' ----------------------------------------------------------------
' Procedure Name:   ShareName
' Purpose:          Returns the network share name for a specified drive.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get ShareName() As String
  ShareName = my_drv.ShareName
End Property


' ----------------------------------------------------------------
' Procedure Name:   TotalSize
' Purpose:          Returns the total space, in bytes, of a drive or network share.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' ----------------------------------------------------------------
Public Property Get TotalSize() As Long
  TotalSize = FormatNumber(my_drv.TotalSize / 1024, 0)
End Property


' ----------------------------------------------------------------
' Procedure Name:   VolumeName
' Purpose:          Returns the volume name of the specified drive. Read/write.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get VolumeName() As String
  VolumeName = my_drv.VolumeName
End Property

' Properties - End