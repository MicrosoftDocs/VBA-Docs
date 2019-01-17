---
title: ChDrive statement (VBA)
keywords: vblr6.chm1008865
f1_keywords:
- vblr6.chm1008865
ms.prod: office
ms.assetid: b07d5925-fba0-9a50-8197-c782fda0bee5
ms.date: 12/03/2018
localization_priority: Normal
---


# ChDrive statement

Changes the current drive.

## Syntax

**ChDrive** _drive_

The required _drive_ [argument](../../Glossary/vbe-glossary.md#argument) is a [string expression](../../Glossary/vbe-glossary.md#string-expression) that specifies an existing drive. If you supply a zero-length string (""), the current drive doesn't change. If the _drive_ argument is a multiple-character string, **ChDrive** uses only the first letter.

On the Macintosh, **ChDrive** changes the current folder to the root folder of the specified drive.

## Example

This example uses the **ChDrive** statement to change the current drive. On the Macintosh, "HD:" is the default drive name, and **ChDrive** would change the current folder to the root folder of the specified drive. The following example assumes the machine actually has a drive named D.


```vb
ChDrive "D" ' Make "D" the current drive. 

```

## See also

- [ChDir statement](chdir-statement.md)
- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]