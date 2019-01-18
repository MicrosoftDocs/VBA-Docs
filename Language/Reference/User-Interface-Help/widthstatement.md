---
title: Width statement (VBA)
keywords: vblr6.chm1009060
f1_keywords:
- vblr6.chm1009060
ms.prod: office
ms.assetid: 655e73fc-c294-5f82-4c1a-59c2ebd71036
ms.date: 12/03/2018
localization_priority: Normal
---


# Width # statement

Assigns an output line width to a file opened by using the **[Open](open-statement.md)** statement.

## Syntax

**Width #**_filenumber_, _width_

<br/>

The **Width #** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
| _filenumber_|Required. Any valid [file number](../../Glossary/vbe-glossary.md#file-number).|
| _width_|Required. [Numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) in the range 0&ndash;255, inclusive, that indicates how many characters appear on a line before a new line is started. If _width_ equals 0, there is no limit to the length of a line. The default value for _width_ is 0.|

## Example

This example uses the **Width #** statement to set the output line width for a file.


```vb
Dim I 
Open "TESTFILE" For Output As #1 ' Open file for output. 
VBA.Width 1, 5 ' Set output line width to 5. 
For I = 0 To 9 ' Loop 10 times. 
 Print #1, Chr(48 + I); ' Prints five characters per line. 
Next I 
Close #1 ' Close file. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]