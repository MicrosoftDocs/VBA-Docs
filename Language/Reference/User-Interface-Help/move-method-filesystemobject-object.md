---
title: Move method (FileSystemObject object)
keywords: vblr6.chm2182006
f1_keywords:
- vblr6.chm2182006
ms.prod: office
ms.assetid: 9191e310-2b92-fd13-f04a-e34ca2743b7e
ms.date: 12/14/2018
localization_priority: Normal
---


# Move method 

Moves a specified file or folder from one location to another.

## Syntax

_object_.**Move** _destination_

<br/>

The **Move** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[File](file-object.md)** or **[Folder](folder-object.md)** object.|
| _destination_|Required. Destination where the file or folder is to be moved. Wildcard characters are not allowed.|

## Remarks

The results of the **Move** method on a **File** or **Folder** are identical to operations performed by using **FileSystemObject.MoveFile** or **FileSystemObject.MoveFolder**. You should note, however, that the alternative methods are capable of moving multiple files or folders.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
