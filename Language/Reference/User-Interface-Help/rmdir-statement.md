---
title: RmDir statement (VBA)
keywords: vblr6.chm1009007
f1_keywords:
- vblr6.chm1009007
ms.prod: office
ms.assetid: 7bc350d2-7d1a-7c8c-95a8-8dbf5c8f7953
ms.date: 12/03/2018
localization_priority: Normal
---


# RmDir statement

Removes an existing directory or folder.

## Syntax

**RmDir**_path_

The required _path_ [argument](../../Glossary/vbe-glossary.md#argument) is a [string expression](../../Glossary/vbe-glossary.md#string-expression) that identifies the directory or folder to be removed. The _path_ may include the drive. If no drive is specified, **RmDir** removes the directory or folder on the current drive.

## Remarks

An error occurs if you try to use **RmDir** on a directory or folder containing files. Use the **[Kill](kill-statement.md)** statement to delete all files before attempting to remove a directory or folder.

## Example

This example uses the **RmDir** statement to remove an existing directory or folder.

```vb
' Assume that MYDIR is an empty directory or folder. 
RmDir "MYDIR" ' Remove MYDIR. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]