Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private my_fso  As Object ' Scripting.FileSystemObject
Private my_fdr  As Object
'


Private Sub Class_Terminate()

  Set my_fso = Nothing
  Set my_fdr = Nothing
  
End Sub


' Properties - Start

' ----------------------------------------------------------------
' Procedure Name:   Attributes
' Purpose:          Returns the attributes of a specified folder.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' ----------------------------------------------------------------
Public Property Get Attributes() As Long
  Attributes = my_fdr.Attributes
End Property


' ----------------------------------------------------------------
' Procedure Name:   DateCreated
' Purpose:          Returns the date and time when a specified folder was created.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Date
' ----------------------------------------------------------------
Public Property Get DateCreated() As Date
  DateCreated = my_fdr.DateCreated
End Property


' ----------------------------------------------------------------
' Procedure Name:   DateLastAccessed
' Purpose:          Returns the date and time when a specified folder was last accessed.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Date' ----------------------------------------------------------------
Public Property Get DateLastAccessed() As Date
  DateLastAccessed = my_fdr.DateLastAccessed
End Property


' ----------------------------------------------------------------
' Procedure Name:   DateLastModified
' Purpose:          Returns the date and time when a specified folder was last modified.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Date
' ----------------------------------------------------------------
Public Property Get DateLastModified() As Date
  DateLastModified = my_fdr.DateLastModified
End Property


' ----------------------------------------------------------------
' Procedure Name:   Drive
' Purpose:          Returns the drive letter of the drive where the specified folder resides.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' ----------------------------------------------------------------
Public Property Get Drive() As Object
  Set Drive = my_fdr.Drive
End Property


' ----------------------------------------------------------------
' Procedure Name:   IsRootFolder
' Purpose:          Returns True if the specified folder is the root folder; False if it is not.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Boolean
' ----------------------------------------------------------------
Public Property Get IsRootFolder() As Boolean
  IsRootFolder = my_fdr.IsRootFolder
End Property


' ----------------------------------------------------------------
' Procedure Name:   Name
' Purpose:          Sets or returns the name of a specified folder.
' Procedure Kind:   Eigenschaft (Let)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter str_Folder (String):  is the new name of the specified object.
' ----------------------------------------------------------------
Public Property Let Name(ByVal str_Folder As String)
  
  Set my_fso = Nothing
  Set my_fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
  
  Set my_fdr = my_fso.GetFolder(str_Folder)
  
End Property

Public Property Get Name() As String
  Name = my_fdr.Name
End Property


' ----------------------------------------------------------------
' Procedure Name:   ParentFolder
' Purpose:          Returns the parent folder of a specified folder.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get ParentFolder() As String
  ParentFolder = my_fdr.ParentFolder
End Property


' ----------------------------------------------------------------
' Procedure Name:   Path
' Purpose:          Returns the path for a specified folder.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get Path() As String
  Path = my_fdr.Path
End Property


' ----------------------------------------------------------------
' Procedure Name:   ShortName
' Purpose:          Returns the short name of a specified folder (the 8.3 naming convention).
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get ShortName() As String
  ShortName = my_fdr.ShortName
End Property


' ----------------------------------------------------------------
' Procedure Name:   ShortPath
' Purpose:          Returns the short path of a specified folder (the 8.3 naming convention).
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get ShortPath() As String
  ShortPath = my_fdr.ShortPath
End Property


' ----------------------------------------------------------------
' Procedure Name:   Typ
' Purpose:          Returns the type of a specified folder.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get Typ() As String
  Typ = my_fdr.Type
End Property

' Properties - End


' Methods - Start


' ----------------------------------------------------------------
' Procedure Name:   Delete
' Purpose:          Deletes a specified folder.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter force (Boolean): Optional. Boolean value that is True if files or folders with the read-only attribute set are to be deleted; False (default) if they are not.
' ----------------------------------------------------------------
Public Function Delete(Optional force As Boolean = False)
  
  my_fso.DeleteFolder my_fdr.Path, force
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   Move
' Purpose:          Moves a specified folder from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter destination (String): Required. Destination where the file or folder is to be moved. Wildcard characters are not allowed.
' ----------------------------------------------------------------
Public Function Move(ByVal destination As String)
  
  my_fso.MoveFolder my_fdr.Path, destination
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   Copy
' Purpose:          Copies a specified folder from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter destination (String): Required. Destination where the file or folder is to be copied. Wildcard characters are not allowed.
' Parameter overwrite (Boolean):   Optional. Boolean value that is True (default) if existing files or folders are to be overwritten; False if they are not.
' ----------------------------------------------------------------
Public Function Copy(ByVal destination As String _
          , Optional ByVal overwrite As Boolean = True)
  
  my_fso.CopyFolder my_fdr.Path, destination
  
End Function

' Methods - End