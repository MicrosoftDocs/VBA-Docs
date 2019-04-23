---
title: AvailableSpace property (Visual Basic for Applications)
keywords: vblr6.chm2181954
f1_keywords:
- vblr6.chm2181954
ms.prod: office
api_name:
- Office.AvailableSpace
ms.assetid: c7a2a011-1b90-7091-4dcb-0149c75a6ee6
ms.date: 12/19/2018
localization_priority: Normal
---


# AvailableSpace property

Returns the amount of space available to a user on the specified drive or network share.

## Syntax

_object_.**AvailableSpace**

The _object_ is always a **[Drive](drive-object.md)** object.

## Remarks

The value returned by the **AvailableSpace** property is typically the same as that returned by the **[FreeSpace](freespace-property.md)** property. Differences may occur between the two values for computer systems that support quotas.

The following code illustrates the use of the **AvailableSpace** property.

```vb
Sub ShowAvailableSpace(drvPath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    s = "Drive " & UCase(drvPath) & " - " 
    s = s & d.VolumeName  & vbCrLf
    s = s & "Available Space: " & FormatNumber(d.AvailableSpace/1024, 0) 
    s = s & " Kbytes"
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]