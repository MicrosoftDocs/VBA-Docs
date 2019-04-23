---
title: ShareName property (Visual Basic for Applications)
keywords: vblr6.chm2181963
f1_keywords:
- vblr6.chm2181963
ms.prod: office
api_name:
- Office.ShareName
ms.assetid: 913ae336-102c-9c1c-4995-9b37aae79b3e
ms.date: 12/19/2018
localization_priority: Normal
---


# ShareName property

Returns the network share name for a specified drive.

## Syntax

_object_.**ShareName**

The _object_ is always a **[Drive](drive-object.md)** object.

## Remarks

If _object_ is not a network drive, the **ShareName** property returns a zero-length string ("").

The following code illustrates the use of the **ShareName** property.

```vb
Sub ShowDriveInfo(drvpath)
    Dim fs, d, s 
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(drvpath)))
    s = "Drive " & d.DriveLetter & ": - " & d.ShareName
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]