Option Compare Database
Option Explicit


Private Function test_01()
  
  Dim TextStream As New my_TextStream_Object
  
  TextStream.CreateTextFile "C:\tmp\mein_test.txt", True, True
  TextStream.WriteLine "zeile1"
  TextStream.Write_Text "__text__"
  TextStream.WriteLine "zeile2"
  TextStream.WriteBlankLines 2
  TextStream.WriteLine "zeile3"
  TextStream.WriteLine Now
  TextStream.Close_Stream
  
  Set TextStream = Nothing
  
End Function

Private Function test_02()
  
  Dim TextStream As New my_TextStream_Object
  
  TextStream.OpenTextFile "C:\tmp\mein_test.txt", ForReading
  
  Debug.Print TextStream.ReadLine
  Debug.Print TextStream.ReadLine
  Debug.Print TextStream.ReadAll
  
  TextStream.Close_Stream
  
  Set TextStream = Nothing
  
End Function

Private Function test_03()
  
  Dim TextStream As New my_TextStream_Object
  
  TextStream.OpenTextFile "C:\tmp\mein_test.txt", ForReading
  
  Do While Not TextStream.AtEndOfStream = True
    Debug.Print TextStream.ReadLine
    TextStream.SkipLine
  Loop
  
  TextStream.Close_Stream
  
  Set TextStream = Nothing
  
End Function