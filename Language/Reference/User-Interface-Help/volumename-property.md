---
title: VolumeName property (Visual Basic for Applications)
keywords: vblr6.chm2181965
f1_keywords:
- vblr6.chm2181965
ms.prod: office
api_name:
- Office.VolumeName
ms.assetid: 8592ae63-f36e-e87a-8286-72419d7781d0
ms.date: 12/19/2018
localization_priority: Normal
---


# VolumeName property

Sets or returns the volume name of the specified drive. Read/write.

## Syntax

_object_.**VolumeName** [= _newname_ ]

<br/>

The **VolumeName** property has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[Drive](drive-object.md)** object.|
| _newname_|Optional. If provided, _newname_ is the new name of the specified _object_.|

## Remarks

The following code illustrates the use of the **VolumeName** property.

```vb
Sub ShowVolumeInfo(drvpath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(drvpath)))
    s = "Drive " & d.DriveLetter & ": - " & d.VolumeName
    MsgBox s
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]