---
title: MkDir statement (VBA)
keywords: vblr6.chm1008975
f1_keywords:
- vblr6.chm1008975
ms.prod: office
ms.assetid: b79fdad3-a1c2-7af3-c679-09d35d4b0d87
ms.date: 12/03/2018
localization_priority: Normal
---


# MkDir statement

Creates a new directory or folder.

## Syntax

**MkDir** _path_

The required _path_ [argument](../../Glossary/vbe-glossary.md#argument) is a [string expression](../../Glossary/vbe-glossary.md#string-expression) that identifies the directory or folder to be created. The _path_ may include the drive. If no drive is specified, **MkDir** creates the new directory or folder on the current drive.

## Example

This example uses the **MkDir** statement to create a directory or folder. If the drive is not specified, the new directory or folder is created on the current drive.


```vb
MkDir "MYDIR" ' Make new directory or folder. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
