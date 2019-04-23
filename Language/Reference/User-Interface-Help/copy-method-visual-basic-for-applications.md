---
title: Copy method (Visual Basic for Applications)
keywords: vblr6.chm2182004
f1_keywords:
- vblr6.chm2182004
ms.prod: office
ms.assetid: 3477c158-643a-5e29-e4c2-b451e8603542
ms.date: 12/14/2018
localization_priority: Normal
---


# Copy method

Copies a specified file or folder from one location to another.

## Syntax

_object_.**Copy** _destination_, [ _overwrite_ ]

<br/>

The **Copy** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[File](file-object.md)** or **[Folder](folder-object.md)** object.|
| _destination_|Required. Destination where the file or folder is to be copied. Wildcard characters are not allowed.|
| _overwrite_|Optional. **Boolean** value that is **True** (default) if existing files or folders are to be overwritten; **False** if they are not.|

## Remarks

The results of the **Copy** method on a **File** or **Folder** are identical to operations performed by using **FileSystemObject.CopyFile** or **FileSystemObject.CopyFolder** where the file or folder referred to by _object_ is passed as an argument. You should note, however, that the alternative methods are capable of copying multiple files or folders.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
