---
title: DeleteFile method (Visual Basic for Applications)
keywords: vblr6.chm2182036
f1_keywords:
- vblr6.chm2182036
ms.prod: office
api_name:
- Office.DeleteFile
ms.assetid: e036b009-4fd9-297a-de24-acc0dbc96c7a
ms.date: 12/14/2018
localization_priority: Priority
---


# DeleteFile method

Deletes a specified file.

## Syntax

_object_.**DeleteFile** _filespec_, [ _force_ ]

<br/>

The **DeleteFile** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _filespec_|Required. The name of the file to delete. The _filespec_ can contain wildcard characters in the last path component.|
| _force_|Optional. **Boolean** value that is **True** if files with the read-only attribute set are to be deleted; **False** (default) if they are not.|

## Remarks

An error occurs if no matching files are found. The **DeleteFile** method stops on the first error it encounters. No attempt is made to roll back or undo any changes that were made before an error occurred.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
