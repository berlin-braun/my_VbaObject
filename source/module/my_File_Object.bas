Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private my_fso As Object ' Scripting.FileSystemObject
Private my_fle As Object
'


Private Sub Class_Terminate()
  
  Set my_fso = Nothing
  Set my_fle = Nothing
  
End Sub


' Properties - Start

' ----------------------------------------------------------------
' Procedure Name:   Attributes
' Purpose:          Sets or returns the attributes of files. Read/write or read-only, depending on the attribute.
' Procedure Kind:   Eigenschaft (Get)
' Procedure Access: Public
' Return Type: Long
' Author:           Thomas Braun
' Date:             30.03.2023
' ----------------------------------------------------------------
Public Property Get Attributes() As Long
  Attributes = my_fle.Attributes
End Property


' ----------------------------------------------------------------
' Procedure Name:   DateCreated
' Purpose:          Returns the date and time that the specified file was created. Read-only.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Return Type:      Date
' Procedure Access: Public
' ----------------------------------------------------------------
Public Property Get DateCreated() As Date
  DateCreated = my_fle.DateCreated
End Property


' ----------------------------------------------------------------
' Procedure Name:   DateLastAccessed
' Purpose:          Returns the date and time that the specified file was last accessed. Read-only.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Return Type:      Date
' Procedure Access: Public
' ----------------------------------------------------------------
Public Property Get DateLastAccessed() As Date
  DateLastAccessed = my_fle.DateLastAccessed
End Property


' ----------------------------------------------------------------
' Procedure Name:   DateLastModified
' Purpose:          Returns the date and time that the specified file was last modified. Read-only.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Return Type:      Date
' Procedure Access: Public
' ----------------------------------------------------------------
Public Property Get DateLastModified() As Date
  DateLastModified = my_fle.DateLastModified
End Property


' ----------------------------------------------------------------
' Procedure Name:   Drive
' Purpose:          Returns the drive letter of the drive on which the specified file resides. Read-only.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Return Type:      Object
' Procedure Access: Public
' ----------------------------------------------------------------
Public Property Get Drive() As Object
  
  Set Drive = my_fle.Drive

End Property


' ----------------------------------------------------------------
' Procedure Name:   Name
' Purpose:          Sets or returns the name of a specified file. Read/write.
' Procedure Kind:   Eigenschaft (Let)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter str_File (String): is the new name of the specified object.
' ----------------------------------------------------------------
Public Property Let Name(ByVal str_File As String)
  
  Set my_fso = Nothing
  Set my_fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
  
  Set my_fle = my_fso.GetFile(str_File)
  
End Property

Public Property Get Name() As String
  Name = my_fle.Name
End Property


' ----------------------------------------------------------------
' Procedure Name:   ParentFolder
' Purpose:          Returns the folder object for the parent of the specified file. Read-only.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' ----------------------------------------------------------------
Public Property Get ParentFolder() As Object
  
  Set ParentFolder = my_fle.ParentFolder

End Property


' ----------------------------------------------------------------
' Procedure Name:   Path
' Purpose:          Returns the path for a specified file.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get Path() As String
  Path = my_fle.Path
End Property


' ----------------------------------------------------------------
' Procedure Name:   ShortName
' Purpose:          Returns the short name used by programs that require the earlier 8.3 naming convention.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get ShortName() As String
  ShortName = my_fle.ShortName
End Property


' ----------------------------------------------------------------
' Procedure Name:   ShortPath
' Purpose:          Returns the short path used by programs that require the earlier 8.3 file naming convention.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get ShortPath() As String
  ShortPath = my_fle.ShortPath
End Property


' ----------------------------------------------------------------
' Procedure Name:   Size
' Purpose:          For files, returns the size, in bytes, of the specified file. For folders, returns the size, in bytes, of all files and subfolders contained in the folder.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Long
' ----------------------------------------------------------------
Public Property Get Size() As Long
  Size = my_fle.Size
End Property


' ----------------------------------------------------------------
' Procedure Name:   Typ
' Purpose:          Returns information about the type of a file or folder. For example, for files ending in .TXT, "Text Document" is returned.
' Procedure Kind:   Eigenschaft (Get)
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      String
' ----------------------------------------------------------------
Public Property Get Typ() As String
  Typ = my_fle.Type
End Property

' Properties - End


' Methods - Start


' ----------------------------------------------------------------
' Procedure Name:   Delete
' Purpose:          Deletes a specified file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter force (Boolean): Optional. Boolean value that is True if files or folders with the read-only attribute set are to be deleted; False (default) if they are not.
' ----------------------------------------------------------------
Public Function Delete(Optional force As Boolean = False)
  
  my_fso.DeleteFile my_fle.Path
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   Move
' Purpose:          Moves a specified file from one location to another.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter destination (String): Required. Destination where the file or folder is to be moved. Wildcard characters are not allowed.
' ----------------------------------------------------------------
Public Function Move(ByVal destination As String)
  
  my_fso.MoveFile my_fle.Path, destination
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   Copy
' Purpose:          Copies a specified file from one location to another.
' Procedure Kind:   Function
' Procedure Access: Public
' Author: Thomas Braun
' Date: 30.03.2023
' Parameter destination (String): Required. Destination where the file or folder is to be copied. Wildcard characters are not allowed.
' Parameter overwrite (Boolean): Optional. Boolean value that is True (default) if existing files or folders are to be overwritten; False if they are not.
' ----------------------------------------------------------------
Public Function Copy(ByVal destination As String, Optional ByVal overwrite As Boolean = True)
  
  my_fso.CopyFile my_fle.Path, destination, overwrite
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   CreateTextFile
' Purpose:          Creates a specified file name and returns a TextStream object that can be used to read from or write to the file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Parameter filename (String): Required. String expression that identifies the file to create.
' Parameter overwrite (Boolean): Optional. Boolean value that indicates if an existing file can be overwritten. The value is True if the file can be overwritten; False if it can't be overwritten. If omitted, existing files can be overwritten.
' Parameter unicode (Boolean): Optional. Boolean value that indicates whether the file is created as a Unicode or ASCII file. The value is True if the file is created as a Unicode file; False if it's created as an ASCII file. If omitted, an ASCII file is assumed.
' ----------------------------------------------------------------
Public Function CreateTextFile(filename As String _
                    , Optional overwrite As Boolean = True _
                    , Optional unicode As Boolean = False)
  Set my_fso = Nothing
  Set my_fle = Nothing
  
  Set my_fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
  Set my_fle = my_fso.CreateTextFile(filename, overwrite, unicode)
  
End Function


' ----------------------------------------------------------------
' Procedure Name:   OpenAsTextStream
' Purpose:          Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.
' Procedure Kind:   Function
' Author:           Thomas Braun
' Date:             30.03.2023
' Procedure Access: Public
' Return Type:      Object
' Parameter filename (String): Name of a File.
' Parameter mode (my_IOMode): Optional. Indicates input/output mode. Can be one of three constants: ForReading, ForWriting, or ForAppending.
' Parameter format (my_Tristate): Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.
' ----------------------------------------------------------------
Public Function OpenAsTextStream(ByVal filename As String _
                      , Optional ByVal mode As my_IOMode = ForAppending _
                      , Optional ByVal format As my_Tristate = TristateUseDefault) As Object
  Dim my_txs As Object

  Set my_fso = Nothing
  Set my_txs = Nothing

  Set my_fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
  Set my_fle = my_fso.GetFile(filename)
  Set my_txs = my_fle.OpenAsTextStream(filename, mode, format)

  Set OpenAsTextStream = my_txs

  Set my_txs = Nothing
  
End Function

' Methods - End