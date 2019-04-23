---
title: GetParentFolderName method (Visual Basic for Applications)
keywords: vblr6.chm2182056
f1_keywords:
- vblr6.chm2182056
ms.prod: office
api_name:
- Office.GetParentFolderName
ms.assetid: 445e969a-6a01-6cb0-aff7-378717277c69
ms.date: 12/14/2018
localization_priority: Normal
---


# GetParentFolderName method

Returns a string containing the name of the parent folder of the last component in a specified path.

## Syntax

_object_.**GetParentFolderName** (_path_)

<br/>

The **GetParentFolderName** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _path_|Required. The path specification for the component whose parent folder name is to be returned.|

## Remarks

The **GetParentFolderName** method returns a zero-length string ("") if there is no parent folder for the component specified in the _path_ argument.

> [!NOTE] 
> The **GetParentFolderName** method works only on the provided _path_ string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]