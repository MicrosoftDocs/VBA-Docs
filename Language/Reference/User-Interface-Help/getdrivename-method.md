---
title: GetDriveName method (Visual Basic for Applications)
keywords: vblr6.chm2182049
f1_keywords:
- vblr6.chm2182049
ms.prod: office
api_name:
- Office.GetDriveName
ms.assetid: cbd31a00-c593-defe-71ad-d1ddde377737
ms.date: 12/14/2018
localization_priority: Normal
---


# GetDriveName method

Returns a string containing the name of the drive for a specified path.

## Syntax

_object_.**GetDriveName** (_path_)

<br/>

The **GetDriveName** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _path_|Required. The path specification for the component whose drive name is to be returned.|

## Remarks

The **GetDriveName** method returns a zero-length string ("") if the drive can't be determined.

> [!NOTE] 
> The **GetDriveName** method works only on the provided _path_ string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]