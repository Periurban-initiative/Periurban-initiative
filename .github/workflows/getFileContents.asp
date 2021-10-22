
<%
'Pass the name of the file to the function.
Function getFileContents(strIncludeFile)
' Set up constants
Const ForReading = 1 
Const Create = False

' Declare local variables
Dim objFSO         ' FileSystemObject
Dim TS             ' TextStreamObject
Dim strLine        ' local variable to store Line
Dim strFileName    ' local variable to store fileName

strFileName = Server.MapPath(strIncludeFile & ".htm")

' Instantiate the FileSystemObject
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

' use Opentextfile Method to Open the text File
Set TS = objFSO.OpenTextFile(strFileName, ForReading, Create)

If Not TS.AtEndOfStream  Then   

 Do While Not TS.AtendOfStream  
 
  strLine = TS.ReadLine ' Read one line at a time
  Response.Write strLine ' Display that line
  
 Loop
 
End If

' close TextStreamObject 
' and destroy local variables to relase memory
TS.Close 
Set TS = Nothing
Set objFSO = Nothing
end function
%>