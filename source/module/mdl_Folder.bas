Option Compare Database
Option Explicit
'
'
' Factory: static properties and methods for class my_Folder_Object
'
'

Public Function folder_INIT(ByVal folder_Name As String) As my_Folder_Object
  
  Dim m_Folder As New my_Folder_Object
  
  m_Folder.Name = folder_Name
  
  Set folder_INIT = m_Folder
  
  Set m_Folder = Nothing
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_Attributes
' Purpose:          Returns the attributes of a specified folder.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_Attributes(ByVal foldername As String) As Long
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_Attributes = m_Folder.Attributes
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_DateCreated
' Purpose:          Returns the date and time when a specified folder was created.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Date
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_DateCreated(ByVal foldername As String) As Date
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_DateCreated = m_Folder.DateCreated
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_DateLastAccessed
' Purpose:          Returns the date and time when a specified folder was last accessed.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Date
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_DateLastAccessed(ByVal foldername As String) As Date
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_DateLastAccessed = m_Folder.DateLastAccessed
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_DateLastModified
' Purpose:          Returns the date and time when a specified folder was last modified.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Date
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_DateLastModified(ByVal foldername As String) As Date
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_DateLastModified = m_Folder.DateLastModified
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_Drive
' Purpose:          Returns the drive letter of the drive where the specified folder resides.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_Drive(ByVal foldername As String) As String
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_Drive = m_Folder.Drive
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_IsRootFolder
' Purpose:          Returns True if the specified folder is the root folder; False if it is not.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Boolean
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_IsRootFolder(ByVal foldername As String) As Boolean
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_IsRootFolder = m_Folder.IsRootFolder
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_ParentFolder
' Purpose:          Returns the parent folder of a specified folder.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_ParentFolder(ByVal foldername As String) As String
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_ParentFolder = m_Folder.ParentFolder
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_Path
' Purpose:          Returns the path for a specified folder.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_Path(ByVal foldername As String) As String
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_Path = m_Folder.Path
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_ShortName
' Purpose:          Returns the short name of a specified folder (the 8.3 naming convention).
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_ShortName(ByVal foldername As String) As String
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_ShortName = m_Folder.ShortName
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_ShortPath
' Purpose:          Returns the short path of a specified folder (the 8.3 naming convention).
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_ShortPath(ByVal foldername As String) As String
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_ShortPath = m_Folder.ShortPath
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_Typ
' Purpose:          Returns the type of a specified folder.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter foldername (String): the specific foldername
' ----------------------------------------------------------------
Public Function folder_Typ(ByVal foldername As String) As String
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  folder_Typ = m_Folder.Typ
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_Delete
' Purpose:          Deletes a specified folder.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter foldername (String): the specific foldername

' Parameter force (Boolean): Optional. Boolean value that is True if folders with the read-only attribute set are to be deleted; False (default) if they are not.
' ----------------------------------------------------------------
Public Function folder_Delete(ByVal foldername As String _
                   , Optional ByVal force As Boolean = False)
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  m_Folder.Delete force
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_Move
' Purpose:          Moves a specified folder from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter foldername (String): the specific foldername

' Parameter destination (String):  Required. Destination where the folder is to be moved. Wildcard characters are not allowed.
' ----------------------------------------------------------------
Public Function folder_Move(ByVal foldername As String _
                          , ByVal destination As String)
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  m_Folder.Move destination
  
  Set m_Folder = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   folder_Copy
' Purpose:          Copies a specified folder from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter foldername (String): the specific foldername
' Parameter destination (String): Required. Destination where the folder is to be copied. Wildcard characters are not allowed.
' Parameter overwrite (Boolean):   Optional. Boolean value that is True (default) if existing folders are to be overwritten; False if they are not.
' ----------------------------------------------------------------
Public Function folder_Copy(ByVal foldername As String _
                          , ByVal destination As String _
                 , Optional ByVal overwrite As Boolean = True)
  Dim m_Folder As New my_Folder_Object
  
  Set m_Folder = folder_INIT(foldername)
  
  m_Folder.Copy destination, overwrite
  
  Set m_Folder = Nothing
End Function