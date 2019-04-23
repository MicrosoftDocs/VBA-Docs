---
title: ParentFolder property (Visual Basic for Applications)
keywords: vblr6.chm2181999
f1_keywords:
- vblr6.chm2181999
ms.prod: office
api_name:
- Office.ParentFolder
ms.assetid: 980e6bf3-fdc2-4335-7587-e5e932aee0a2
ms.date: 12/19/2018
localization_priority: Normal
---


# ParentFolder property

Returns the folder object for the parent of the specified file or folder. Read-only.

## Syntax

_object_.**ParentFolder**

The _object_ is always a **[File](file-object.md)** or **[Folder](folder-object.md)** object.

## Remarks

The following code illustrates the use of the **ParentFolder** property with a file.

```vb
Sub ShowFileAccessInfo(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = UCase(f.Name) & " in " & UCase(f.ParentFolder) & vbCrLf
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