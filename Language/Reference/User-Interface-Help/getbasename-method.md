---
title: GetBaseName method (Visual Basic for Applications)
keywords: vblr6.chm2182047
f1_keywords:
- vblr6.chm2182047
ms.prod: office
api_name:
- Office.GetBaseName
ms.assetid: 2f3af3ff-a996-e2f7-0048-1f5aa891d674
ms.date: 12/14/2018
localization_priority: Priority
---


# GetBaseName method

Returns a string containing the base name of the last component, less any file extension, in a path.

## Syntax

_object_.**GetBaseName** (_path_)

<br/>

The **GetBaseName** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _path_|Required. The path specification for the component whose base name is to be returned.|

## Remarks

The **GetBaseName** method returns a zero-length string ("") if no component matches the _path_ argument.

> [!NOTE] 
> The **GetBaseName** method works only on the provided _path_ string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]