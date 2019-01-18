---
title: GetExtensionName method (Visual Basic for Applications)
keywords: vblr6.chm2182052
f1_keywords:
- vblr6.chm2182052
ms.prod: office
api_name:
- Office.GetExtensionName
ms.assetid: 0fa9da71-7938-c50c-6fed-8a23d6a680d1
ms.date: 12/14/2018
localization_priority: Normal
---


# GetExtensionName method

Returns a string containing the extension name for the last component in a path.

## Syntax

_object_.**GetExtensionName** (_path_)

<br/>

The **GetExtensionName** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _path_|Required. The path specification for the component whose extension name is to be returned.|

## Remarks

For network drives, the root directory (**\**) is considered to be a component.

The **GetExtensionName** method returns a zero-length string ("") if no component matches the _path_ argument.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]