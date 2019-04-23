---
title: GetFile method (Visual Basic for Applications)
keywords: vblr6.chm2182054
f1_keywords:
- vblr6.chm2182054
ms.prod: office
api_name:
- Office.GetFile
ms.assetid: bdb2737e-7836-4dac-9216-6f1bd8f92aa8
ms.date: 12/14/2018
localization_priority: Normal
---


# GetFile method

Returns a **[File](file-object.md)** object corresponding to the file in a specified path.

## Syntax

_object_.**GetFile** (_filespec_)

<br/>

The **GetFile** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _filespec_|Required. The _filespec_ is the path (absolute or relative) to a specific file.|

## Remarks

An error occurs if the specified file does not exist.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
