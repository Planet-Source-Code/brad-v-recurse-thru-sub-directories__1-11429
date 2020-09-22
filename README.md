<div align="center">

## Recurse thru sub directories


</div>

### Description

I haven't seen this on PSC using the filesystem object or a collection. very neat and fast since you don't need to go through it twice to redim an array. the only slow down with this code is the print statement.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brad V](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brad-v.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brad-v-recurse-thru-sub-directories__1-11429/archive/master.zip)

### API Declarations

```
Dim colDirs As New Collection
Dim objFso As New FileSystemObject
'make sure under Project|References
'you have Microsoft scripting Runtime checked
```


### Source Code

```
Private Sub Command1_Click()
Dim lForIndex As Long
  Set colDirs = Nothing
  Set colDirs = New Collection
  Me.List1.Clear
  DoEvents
  colToFill.Add Item:=endInSlash("C:")
  Call makeTree("C:", colDirs)
  For lForIndex = 1 To colDirs.Count
    Debug.Print colDirs.Item(lForIndex)
  Next lForIndex
End Sub
Sub makeTree(ByVal inPath As String, ByRef colToFill As Collection)
Dim objDir1 As Folder
Dim objDir2 As Folder
Dim sCurrentDir As String
  sCurrentDir = endInSlash(inPath)
  Set objDir1 = objFso.GetFolder(sCurrentDir)
  For Each objDir2 In objDir1.SubFolders
    colToFill.Add Item:=sCurrentDir & objDir2.Name
    Call makeTree(sCurrentDir & objDir2.Name, colToFill)
  Next objDir2
  Set objDir1 = Nothing
  Set objDir2 = Nothing
End Sub
Function endInSlash(ByVal inString As String) As String
  If Right$(inString, 1) <> "\" Then
    endInSlash = inString & "\"
  Else
    endInSlash = inString
  End If
End Function
```

