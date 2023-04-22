Dim Group,User,objFSO,objFolder,source,target,StrDomain


Set objFSO=CreateObject("Scripting.FileSystemObject")



source="C:\Users\Alexandru Andercou\Desktop\PIU_PRJ\Andercou_Alexandru\react-demo\src"
target="C:\Users\Alexandru Andercou\Desktop\PIU_PRJ\Andercou_Alexandru\componente"'
tt="C:\Users\Alexandru Andercou\Desktop\test_copy"
Set objFolder = objFSO.GetFolder(source)


dim list
Set list = CreateObject("System.Collections.ArrayList")
list.Add "AddTutorials"
list.Add "commons"
list.Add "FilterBox"
list.Add "VolunteerAplications"
list.Add "LessonImage"
list.Add "FirstAidLessons"
list.Add "MockUpApplicationVerification"


For Each file in objFolder.SubFolders
 If  list.Contains(file.name)   Then
   Msgbox file.name
   file.Copy target & "\" & file.name
End If
Next