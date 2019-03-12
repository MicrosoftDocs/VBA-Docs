---
title: GetFolder method (Visual Basic for Applications)
keywords: vblr6.chm2182055
f1_keywords:
- vblr6.chm2182055
ms.prod: office
api_name:
- Office.GetFolder
ms.assetid: 772f1ae7-ac29-d4b4-e08a-d8553375510d
ms.date: 12/14/2018
localization_priority: Normal
---


# GetFolder method

Returns a **[Folder](folder-object.md)** object corresponding to the folder in a specified path.

## Syntax

_object_.**GetFolder** (_folderspec_)

<br/>

The **GetFolder** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _folderspec_|Required. The _folderspec_ is the path (absolute or relative) to a specific folder.|

## Remarks

An error occurs if the specified folder does not exist.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
