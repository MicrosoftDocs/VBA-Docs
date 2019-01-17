---
title: GetDrive method (Visual Basic for Applications)
keywords: vblr6.chm2182048
f1_keywords:
- vblr6.chm2182048
ms.prod: office
api_name:
- Office.GetDrive
ms.assetid: bd11dc26-b806-864c-b30b-6c74b7701901
ms.date: 12/14/2018
localization_priority: Normal
---


# GetDrive method

Returns a **Drive** object corresponding to the drive in a specified path.

## Syntax

_object_.**GetDrive** (_drivespec_)

<br/>

The **GetDrive** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _drivespec_|Required. The _drivespec_ argument can be a drive letter (c), a drive letter with a colon appended (c:), a drive letter with a colon and path separator appended (c:\), or any network share specification (\\computer2\share1).|

## Remarks

For network shares, a check is made to ensure that the share exists.

An error occurs if _drivespec_ does not conform to one of the accepted forms or does not exist.

To call the **GetDrive** method on a normal path string, use the following sequence to get a string that is suitable for use as _drivespec_:

```vb
DriveSpec = GetDriveName(GetAbsolutePathName(Path))

```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]