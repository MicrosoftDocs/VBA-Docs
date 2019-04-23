---
title: FileSystem property (Visual Basic for Applications)
keywords: vblr6.chm2181957
f1_keywords:
- vblr6.chm2181957
ms.prod: office
api_name:
- Office.FileSystem
ms.assetid: 123ba29e-0b94-0afe-5f3d-323e903dd38e
ms.date: 12/19/2018
localization_priority: Normal
---


# FileSystem property

Returns the type of file system in use for the specified drive.

## Syntax

_object_.**FileSystem**

The _object_ is always a **[Drive](drive-object.md)** object.

## Remarks

Available return types include FAT, NTFS, and CDFS.

The following code illustrates the use of the **FileSystem** property.

```vb
Sub ShowFileSystemType
    Dim fs,d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive("e:")
    s = d.FileSystem
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]