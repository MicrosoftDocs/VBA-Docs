---
title: Reset statement (VBA)
keywords: vblr6.chm1009002
f1_keywords:
- vblr6.chm1009002
ms.prod: office
ms.assetid: 7fb7dedd-dcfd-08a1-37e1-fde804b267e4
ms.date: 12/03/2018
localization_priority: Normal
---


# Reset statement

Closes all disk files opened by using the **[Open](open-statement.md)** statement.

## Syntax

**Reset**

## Remarks

The **Reset** statement closes all active files opened by the **Open** statement and writes the contents of all file buffers to disk.

## Example

This example uses the **Reset** statement to close all open files and write the contents of all file buffers to disk. Note the use of the **Variant** variable as both a string and a number.


```vb
Dim FileNumber 
For FileNumber = 1 To 5 ' Loop 5 times. 
 ' Open file for output. FileNumber is concatenated into the string 
 ' TEST for the file name, but is a number following a #. 
 Open "TEST" & FileNumber For Output As #FileNumber 
 Write #FileNumber, "Hello World" ' Write data to file. 
Next FileNumber 
Reset ' Close files and write contents 
 ' to disk. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]