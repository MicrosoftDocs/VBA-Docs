---
title: GetFileName method (Visual Basic for Applications)
keywords: vblr6.chm2182053
f1_keywords:
- vblr6.chm2182053
ms.prod: office
ms.assetid: af5ca68f-ec3e-409c-dcb4-75202169ccb8
ms.date: 12/14/2018
---


# GetFileName method

Returns the last component of a specified path that is not part of the drive specification.

## Syntax

_object_.**GetFileName** (_pathspec_)

<br/>

The **GetFileName** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **[FileSystemObject](filesystemobject-object.md)**.|
| _pathspec_|Required. The path (absolute or relative) to a specific file.|

## Remarks

The **GetFileName** method returns a zero-length string ("") if _pathspec_ does not end with the named component.

> [!NOTE] 
> The **GetFileName** method works only on the provided path string. It does not attempt to resolve the path, nor does it check for the existence of the specified path.

## See also

- [Methods (Visual Basic for Applications)](../methods-visual-basic-for-applications.md)
