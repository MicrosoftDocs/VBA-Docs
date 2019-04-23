---
title: Path property (FileSystemObject object)
keywords: vblr6.chm2181960
f1_keywords:
- vblr6.chm2181960
ms.prod: office
ms.assetid: 15eed13b-9252-e195-0c54-9e3c82ce987f
ms.date: 12/19/2018
localization_priority: Normal
---


# Path property 

Returns the path for a specified file, folder, or drive.

## Syntax

_object_.**Path**

The _object_ is always a **[File](file-object.md)**, **[Folder](folder-object.md)**, or **[Drive](drive-object.md)** object.

## Remarks

For drive letters, the root drive is not included. For example, the path for the C drive is `C:`, not `C:\`.

The following code illustrates the use of the **Path** property with a **File** object.

```vb
Sub ShowFileAccessInfo(filespec)
    Dim fs, d, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = UCase(f.Path) & vbCrLf
    s = s & "Created: " & f.DateCreated & vbCrLf
    s = s & "Last Accessed: " & f.DateLastAccessed & vbCrLf
    s = s & "Last Modified: " & f.DateLastModified  
    MsgBox s, 0, "File Access Info"
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
