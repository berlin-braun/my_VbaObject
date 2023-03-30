Option Compare Database
Option Explicit
'
'
' Factory: static properties and methods for class my_File_Object
'
'


Public Function file_INIT(ByVal file_Name As String) As my_File_Object
  
  Dim m_File As New my_File_Object
  
  m_File.Name = file_Name
  
  Set file_INIT = m_File
  
  Set m_File = Nothing
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_Attributes
' Purpose:          returns the attributes of files. Read/write or read-only, depending on the attribute.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_Attributes(ByVal filename As String) As Long
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_Attributes = m_File.Attributes
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_DateCreated
' Purpose:          Returns the date and time that the specified file was created. Read-only.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Date
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_DateCreated(ByVal filename As String) As Date
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_DateCreated = m_File.DateCreated

  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_DateLastAccessed
' Purpose:          Returns the date and time that the specified file was last accessed. Read-only.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Date
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_DateLastAccessed(ByVal filename As String) As Date
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_DateLastAccessed = m_File.DateLastAccessed
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_DateLastModified
' Purpose:          Returns the date and time that the specified file was last modified. Read-only.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Date
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_DateLastModified(ByVal filename As String) As Date
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_DateLastModified = m_File.DateLastModified
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_Drive
' Purpose:          Returns the drive letter of the drive on which the specified file resides. Read-only.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_Drive(ByVal filename As String) As Object
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  Set file_Drive = m_File.Drive
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_ParentFolder
' Purpose:          Returns the folder object for the parent of the specified file. Read-only.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_ParentFolder(ByVal filename As String) As Object
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  Set file_ParentFolder = m_File.ParentFolder
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_Path
' Purpose:          Returns the path for a specified file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_Path(ByVal filename As String) As String
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_Path = m_File.Path
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_ShortName
' Purpose:          Returns the short name used by programs that require the earlier 8.3 naming convention.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_ShortName(ByVal filename As String) As String
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_ShortName = m_File.ShortName
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_ShortPath
' Purpose:          Returns the short path used by programs that require the earlier 8.3 file naming convention.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_ShortPath(ByVal filename As String) As String
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_ShortPath = m_File.ShortPath
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_Size
' Purpose:          Returns the size, in bytes, of the specified file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_Size(ByVal filename As String) As Long
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_Size = m_File.Size
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_Typ
' Purpose:          Returns information about the type of a file. For example, for files ending in .TXT, "Text Document" is returned.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' Parameter filename (String): Name of the specific file.
' ----------------------------------------------------------------
Public Function file_Typ(ByVal filename As String) As String
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_Typ = m_File.Typ
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_Delete
' Purpose:          Deletes a specified file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter filename (String): Name of the specific file.

' Parameter force (Boolean): Optional. Boolean value that is True if files with the read-only attribute set are to be deleted; False (default) if they are not.
' ----------------------------------------------------------------
Public Function file_Delete(ByVal filename As String _
                 , Optional ByVal force As Boolean = False)
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  m_File.Delete force
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_Move
' Purpose:          Moves a specified file from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter filename (String): Name of the specific file.

' Parameter destination (String): Required. Destination where the file or folder is to be moved. Wildcard characters are not allowed.
' ----------------------------------------------------------------
Public Function file_Move(ByVal filename As String _
                        , ByVal destination As String)
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  m_File.Move destination
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_Copy
' Purpose:          Copies a specified file from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter filename (String): Name of the specific file.
' Parameter destination (String): Required. Destination where the file is to be copied. Wildcard characters are not allowed.
' Parameter overwrite (Boolean): Optional. Boolean value that is True (default) if existing files or folders are to be overwritten; False if they are not.
' ----------------------------------------------------------------
Public Function file_Copy(ByVal filename As String _
                        , ByVal destination As String _
               , Optional ByVal overwrite As Boolean)
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  m_File.Copy destination, overwrite
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_CreateTextFile
' Purpose:          Creates a specified file name and returns a TextStream object that can be used to read from or write to the file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter filename (String): Name of the specific file.
' Parameter overwrite (Boolean): Optional. Boolean value that indicates if an existing file can be overwritten. The value is True if the file can be overwritten; False if it can't be overwritten. If omitted, existing files can be overwritten.
' Parameter unicode (Boolean): Optional. Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is True if the file is created as a Unicode file; False if it's created as an ASCII file. If omitted, an ASCII file is assumed.
' ----------------------------------------------------------------
Public Function file_CreateTextFile(ByVal filename As String _
                         , Optional ByVal overwrite As Boolean = True _
                         , Optional ByVal unicode As Boolean = False)
  Dim m_File    As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  Set file_CreateTextFile = m_File.CreateTextFile(filename, overwrite, unicode)
  
  Set m_File = Nothing
End Function


' ----------------------------------------------------------------
' Procedure Name:   file_OpenAsTextStream
' Purpose:          Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' Parameter filename (String): Name of the specific file.
' Parameter mode (my_IOMode): Optional. Indicates input/output mode. Can be one of three constants: ForReading, ForWriting, or ForAppending.
' Parameter format (my_Tristate): Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.
' ----------------------------------------------------------------
Public Function file_OpenAsTextStream(ByVal filename As String _
                           , Optional ByVal mode As my_IOMode = ForAppending _
                           , Optional ByVal format As my_Tristate = TristateUseDefault) As Object
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  Set file_OpenAsTextStream = m_File.OpenAsTextStream(filename, mode, format)
  
  Set m_File = Nothing
End Function