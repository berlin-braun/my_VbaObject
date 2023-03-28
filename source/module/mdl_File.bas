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


Public Function file_Attributes(ByVal filename As String) As Long
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_Attributes = m_File.Attributes
  
  Set m_File = Nothing
End Function

Public Function file_DateCreated(ByVal filename As String) As Date
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_DateCreated = m_File.DateCreated

  Set m_File = Nothing
End Function

Public Function file_DateLastAccessed(ByVal filename As String) As Date
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_DateLastAccessed = m_File.DateLastAccessed
  
  Set m_File = Nothing
End Function

Public Function file_DateLastModified(ByVal filename As String) As Date
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_DateLastModified = m_File.DateLastModified
  
  Set m_File = Nothing
End Function

Public Function file_Drive(ByVal filename As String) As Object
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  Set file_Drive = m_File.Drive
  
  Set m_File = Nothing
End Function

Public Function file_ParentFolder(ByVal filename As String) As Object
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  Set file_ParentFolder = m_File.ParentFolder
  
  Set m_File = Nothing
End Function

Public Property Get file_Path(ByVal filename As String) As String
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_Path = m_File.Path
  
  Set m_File = Nothing
End Property

Public Function file_ShortName(ByVal filename As String) As String
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_ShortName = m_File.ShortName
  
  Set m_File = Nothing
End Function

Public Function file_ShortPath(ByVal filename As String) As String
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_ShortPath = m_File.ShortPath
  
  Set m_File = Nothing
End Function

Public Function file_Size(ByVal filename As String) As Long
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_Size = m_File.Size
  
  Set m_File = Nothing
End Function

Public Function file_Typ(ByVal filename As String) As String
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  file_Typ = m_File.Typ
  
  Set m_File = Nothing
End Function


Public Function file_Delete(ByVal filename As String _
                 , Optional ByVal force As Boolean = False)
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  m_File.Delete force
  
  Set m_File = Nothing
End Function

Public Function file_Move(ByVal filename As String _
                        , ByVal destination As String)
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  m_File.Move destination
  
  Set m_File = Nothing
End Function

Public Function file_Copy(ByVal filename As String _
                        , ByVal destination As String)
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  m_File.Copy destination
  
  Set m_File = Nothing
End Function

Public Function file_CreateTextFile(filename As String _
                         , Optional overwrite As Boolean = True _
                         , Optional unicode As Boolean = False)
  Dim m_File    As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  Set file_CreateTextFile = m_File.CreateTextFile(filename, overwrite, unicode)
  
  Set m_File = Nothing
End Function

Public Function file_OpenAsTextStream(ByVal filename As String _
                           , Optional ByVal mode As IOMode = ForAppending _
                           , Optional ByVal format As Tristate = TristateUseDefault) As Object
  Dim m_File As New my_File_Object
  
  Set m_File = file_INIT(filename)
  
  Set file_OpenAsTextStream = m_File.OpenAsTextStream(filename, mode, format)
  
  Set m_File = Nothing
End Function