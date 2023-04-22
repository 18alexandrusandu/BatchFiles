
Function BrowseForFile() 
       
    Dim wso: Set wso = CreateObject("Wscript.Shell")
    Dim hta: hta = """about:<input type=file id=file>"                          & vbcrlf _
                 & "<script>resizeTo(0,0);"                                     & vbCrLf _
                 & "   file.click();"                                           & VbCrLf _
                 & "   fso = new ActiveXObject('Scripting.FileSystemObject');"  & vbCrLf _
                 & "   fso.GetStandardStream(1).WriteLine(file.value);"         & vbCrLf _
                 & "   close();"                                                & VbCrLf _
                 & "</script>"""

     Dim value, value2,value3
     Set value=wso.Exec("mshta.exe " & hta)
     Set value2=value.StdOut
     value3=value2.ReadAll()
     Msgbox value3
     Return "magie"
      
  
End Function

Dim objFSO,objFolder,source,target,constraintsFile,line,file,list,fso


Set objFSO=CreateObject("Scripting.FileSystemObject")
Set fso = CreateObject("Scripting.FileSystemObject")


source=InputBox("Source folder")
target=InputBox("Target folder")


f2 = BrowseForFile() 

constraintsFile= f2
Msgbox "file:" & constraintFile


Set list = CreateObject("System.Collections.ArrayList")


Set file = fso.OpenTextFile(constraintsFile, ForReading)
Do Until file.AtEndOfStream
  line = file.Readline
  list.Add line 
Loop


Set objFolder = objFSO.GetFolder(source)

For Each file in objFolder.SubFolders
 If  list.Contains(file.name)   Then
   Msgbox file.name
   file.Copy target & "\" & file.name
End If
Next