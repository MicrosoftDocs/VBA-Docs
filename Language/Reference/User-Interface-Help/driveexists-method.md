---
title: DriveExists method (Visual Basic for Applications)
keywords: vblr6.chm2182038
f1_keywords:
- vblr6.chm2182038
ms.prod: office
api_name:
- Office.DriveExists
ms.assetid: ddba70e5-8b60-4ce6-631f-fb10f81a6d93
ms.date: 12/14/2018
localization_priority: Normal
---


# DriveExists method

Returns **True** if the specified drive exists; **False** if it does not.

## Syntax

_object_.**DriveExists** (_drivespec_)

<br/>

The **DriveExists** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _drivespec_|Required. A drive letter or a path specification for the root of the drive.|

## Remarks

For drives with removable media, the **DriveExists** method returns **True** even if there are no media present. Use the **IsReady** property of the **[Drive](drive-object.md)** object to determine if a drive is ready.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]