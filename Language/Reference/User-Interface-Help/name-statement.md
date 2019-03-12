---
title: Name statement (VBA)
keywords: vblr6.chm1008979
f1_keywords:
- vblr6.chm1008979
ms.prod: office
ms.assetid: c248e962-1265-b871-3ef7-36effb070d2b
ms.date: 12/03/2018
localization_priority: Normal
---


# Name statement

Renames a disk file, directory, or folder.

## Syntax

**Name** _oldpathname_ **As** _newpathname_

<br/>

The **Name** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
| _oldpathname_|Required. [String expression](../../Glossary/vbe-glossary.md#string-expression) that specifies the existing file name and location; may include directory or folder, and drive.|
| _newpathname_|Required. String expression that specifies the new file name and location; may include directory or folder, and drive. The file name specified by _newpathname_ can't already exist.|

## Remarks

The **Name** statement renames a file and moves it to a different directory or folder, if necessary. **Name** can move a file across drives, but it can only rename an existing directory or folder when both _newpathname_ and _oldpathname_ are located on the same drive. **Name** cannot create a new file, directory, or folder.

Using **Name** on an open file produces an error. You must close an open file before renaming it. **Name** [arguments](../../Glossary/vbe-glossary.md#argument) cannot include multiple-character (**\***) and single-character (**?**) wildcards.

## Example

This example uses the **Name** statement to rename a file. For purposes of this example, assume that the directories or folders that are specified already exist. On the Macintosh, "HD:" is the default drive name, and portions of the pathname are separated by colons instead of backslashes.


```vb
Dim OldName, NewName 
OldName = "OLDFILE": NewName = "NEWFILE" ' Define file names. 
Name OldName As NewName ' Rename file. 
 
OldName = "C:\MYDIR\OLDFILE": NewName = "C:\YOURDIR\NEWFILE" 
Name OldName As NewName ' Move and rename file. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
