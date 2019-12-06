---
title: BuildPath method (Visual Basic for Applications)
keywords: vblr6.chm2182031
f1_keywords:
- vblr6.chm2182031
ms.prod: office
api_name:
- Office.BuildPath
ms.assetid: 55f3dbad-0e0a-1968-a749-fe87986e9690
ms.date: 12/14/2018
localization_priority: Normal
---


# BuildPath method

Combines a folder path and the name of a folder or file and returns the combination with valid path separators.

## Syntax

_object_.**BuildPath** (_path_, _name_)

<br/>

The **BuildPath** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _path_|Required. Existing path with which _name_ is combined. Path can be absolute or relative and need not specify an existing folder.|
| _name_|Required. Name of a folder or file being appended to the existing _path_.|

## Remarks

The **BuildPath** method inserts an additional path separator between the existing path and the name, only if necessary.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
