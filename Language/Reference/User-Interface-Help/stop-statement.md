---
title: Stop statement (VBA)
keywords: vblr6.chm1009033
f1_keywords:
- vblr6.chm1009033
ms.prod: office
ms.assetid: 9b6b5394-9b19-8f18-216c-ac64b165218f
ms.date: 12/03/2018
localization_priority: Normal
---


# Stop statement

Suspends execution.

## Syntax

**Stop**

## Remarks

You can place **Stop** statements anywhere in [procedures](../../Glossary/vbe-glossary.md#procedure) to suspend execution. Using the **Stop** statement is similar to setting a [breakpoint](../../Glossary/vbe-glossary.md#breakpoint) in the code.

The **Stop** statement suspends execution, but unlike **End**, it doesn't close any files or clear [variables](../../Glossary/vbe-glossary.md#variable), unless it is in a compiled executable (.exe) file.

## Example

This example uses the **Stop** statement to suspend execution for each iteration through the **For...Next** loop.

```vb
Dim i As Long 
For i = 1 To 10 ' Start For...Next loop. 
 Debug.Print i ' Print i to the Immediate window. 
 Stop ' Stop during each iteration. 
Next i 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]