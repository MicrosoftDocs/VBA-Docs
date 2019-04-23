---
title: Drive property (Visual Basic for Applications)
keywords: vblr6.chm2181976
f1_keywords:
- vblr6.chm2181976
ms.prod: office
api_name:
- Office.Drive
ms.assetid: 34512359-067f-f625-5f19-db7b0faa0138
ms.date: 12/19/2018
localization_priority: Normal
---


# Drive property

Returns the drive letter of the drive on which the specified file or folder resides. Read-only.

## Syntax

_object_.**Drive**

The _object_ is always a **[File](file-object.md)** or **[Folder](folder-object.md)** object.

## Remarks

The following code illustrates the use of the **Drive** property.

```vb
Sub ShowFileAccessInfo(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = f.Name & " on Drive " & UCase(f.Drive) & vbCrLf
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