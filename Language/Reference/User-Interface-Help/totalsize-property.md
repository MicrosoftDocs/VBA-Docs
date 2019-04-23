---
title: TotalSize property (Visual Basic for Applications)
keywords: vblr6.chm2181964
f1_keywords:
- vblr6.chm2181964
ms.prod: office
api_name:
- Office.TotalSize
ms.assetid: 3c5d7904-3abe-2733-abe2-f329979863da
ms.date: 12/19/2018
localization_priority: Normal
---


# TotalSize property

Returns the total space, in bytes, of a drive or network share.

## Syntax

_object_.**TotalSize**

The _object_ is always a **[Drive](drive-object.md)** object.

## Remarks

The following code illustrates the use of the **TotalSize** property.

```vb
Sub ShowSpaceInfo(drvpath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(drvpath)))
    s = "Drive " & d.DriveLetter & ":"
    s = s & vbCrLf
    s = s & "Total Size: " & FormatNumber(d.TotalSize/1024, 0) & " Kbytes"
    s = s & vbCrLf
    s = s & "Available: " & FormatNumber(d.AvailableSpace/1024, 0) & " Kbytes"
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]