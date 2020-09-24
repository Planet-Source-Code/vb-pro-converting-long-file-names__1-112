<div align="center">

## Converting long file names


</div>

### Description

VB4's commands for dealing with file names (such as KILL, MKDIR, and FILECOPY) support long file names without programmer interaction. A number of the Win95 API functions will return only the short name, and you'll notice a number of short file name entries if you're digging through the registration database. Therefore, occasionally you'll need to convert a short file name into a long file name.

This function lets you pass a long file name with no ill effects. The file must exist for the conversion to succeed. Because this routine uses Dir$ and "walks" the path name to do its work, it will not impress you with its speed:
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Pro](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-pro.md)
**Level**          |Unknown
**User Rating**    |4.2 (165 globes from 39 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-pro-converting-long-file-names__1-112/archive/master.zip)





### Source Code

```
Function sLongName(sShortName As String) As String
'sShortName - the provided file name,
'fully qualified, this would usually be
'a short file name, but can be a long file name
'or any combination of long / short parts
'RETURNS: the complete long file name,
'or "" if an error occurs
'an error would usually indicate
'that the file doesn't exist
Dim sTemp As String
Dim sNew As String
Dim iHasBS As Integer
Dim iBS As Integer
If Len(sShortName) = 0 Then Exit Function
sTemp = sShortName
If Right$(sTemp, 1) = "\" Then
sTemp = Left$(sTemp, Len(sTemp) - 1)
iHasBS = True
End If
On Error GoTo MSGLFNnofile
If InStr(sTemp, "\") Then
sNew = ""
Do While InStr(sTemp, "\")
If Len(sNew) Then
sNew = Dir$(sTemp, 54) & "\" & sNew
Else
sNew = Dir$(sTemp, 54)
If sNew = "" Then
sLongName = sShortName
Exit Function
End If
End If
On Error Resume Next
For iBS = Len(sTemp) To 1 Step -1
If ("\" = Mid$(sTemp, iBS, 1)) Then
'found it
Exit For
End If
Next iBS
sTemp = Left$(sTemp, iBS - 1)
Loop
sNew = sTemp & "\" & sNew
Else
sNew = Dir$(sTemp, 54)
End If
MSGLFNresume:
If iHasBS Then
sNew = sNew & "\"
End If
sLongName = sNew
Exit Function
MSGLFNnofile:
sNew = ""
Resume MSGLFNresume
End Function
```

