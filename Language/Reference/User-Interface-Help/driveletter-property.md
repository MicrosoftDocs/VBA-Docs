---
title: DriveLetter property (Visual Basic for Applications)
keywords: vblr6.chm2181955
f1_keywords:
- vblr6.chm2181955
ms.prod: office
api_name:
- Office.DriveLetter
ms.assetid: 29bf179a-8bf7-56de-3cf5-53fd0d2151e0
ms.date: 12/19/2018
localization_priority: Normal
---


# DriveLetter property

Returns the drive letter of a physical local drive or a network share. Read-only.

## Syntax

_object_.**DriveLetter**

The _object_ is always a **[Drive](drive-object.md)** object.

## Remarks

The **DriveLetter** property returns a zero-length string ("") if the specified drive is not associated with a drive letter, for example, a network share that has not been mapped to a drive letter.

The following code illustrates the use of the **DriveLetter** property.

```vb
Sub ShowDriveLetter(drvPath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    s = "Drive " & d.DriveLetter & ": - " 
    s = s & d.VolumeName  & vbCrLf
    s = s & "Free Space: " & FormatNumber(d.FreeSpace/1024, 0) 
    s = s & " Kbytes"
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]