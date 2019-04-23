---
title: Close statement (VBA)
keywords: vblr6.chm1008872
f1_keywords:
- vblr6.chm1008872
ms.prod: office
ms.assetid: a3c4baf2-36a0-2ae9-c7d5-88d836f65d47
ms.date: 12/03/2018
localization_priority: Normal
---


# Close statement

Concludes input/output (I/O) to a file opened by using the **[Open](open-statement.md)** statement.

## Syntax

**Close** [ _filenumberlist_ ]

The optional _filenumberlist_ [argument](../../Glossary/vbe-glossary.md#argument) can be one or more [file numbers](../../Glossary/vbe-glossary.md#file-number) that use the following syntax, where _filenumber_ is any valid file number:
[[ **#** ] _filenumber_ ] [ **,** [ **#** ] _filenumber_ ] **. . .**

## Remarks

If you omit _filenumberlist_, all active files opened by the **Open** statement are closed.

When you close files that were opened for **Output** or **Append**, the final buffer of output is written to the operating system buffer for that file. All buffer space associated with the closed file is released.

When the **Close** statement is executed, the association of a file with its file number ends.

## Example

This example uses the **Close** statement to close all three files opened for **Output**.


```vb
Dim I, FileName 
For I = 1 To 3 ' Loop 3 times. 
 FileName = "TEST" & I ' Create file name. 
 Open FileName For Output As #I ' Open file. 
 Print #I, "This is a test." ' Write string to file. 
Next I 
Close ' Close all 3 open files. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
