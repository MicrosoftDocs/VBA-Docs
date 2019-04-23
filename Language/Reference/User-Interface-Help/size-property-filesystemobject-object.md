---
title: Size property (FileSystemObject object)
keywords: vblr6.chm2182000
f1_keywords:
- vblr6.chm2182000
ms.prod: office
ms.assetid: 8ddecf14-adda-70bd-4d96-42ac0fa18745
ms.date: 12/19/2018
localization_priority: Normal
---


# Size property

For files, returns the size, in bytes, of the specified file. For folders, returns the size, in bytes, of all files and subfolders contained in the folder.

## Syntax

_object_.**Size**

The _object_ is always a **[File](file-object.md)** or **[Folder](folder-object.md)** object.

## Remarks

The following code illustrates the use of the **Size** property with a **Folder** object.

```vb
Sub ShowFolderSize(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(filespec)
    s = UCase(f.Name) & " uses " & f.size & " bytes."
    MsgBox s, 0, "Folder Size Info"
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
