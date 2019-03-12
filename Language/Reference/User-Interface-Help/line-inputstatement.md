---
title: Line Input statement (VBA)
keywords: vblr6.chm1008962
f1_keywords:
- vblr6.chm1008962
ms.prod: office
ms.assetid: 30cfc57e-0d28-b53e-c5cd-0ed99957e25d
ms.date: 12/03/2018
localization_priority: Normal
---


# Line Input # statement

Reads a single line from an open sequential file and assigns it to a [String](../../Glossary/vbe-glossary.md#string-data-type) [variable](../../Glossary/vbe-glossary.md#variable).

## Syntax

**Line Input** #_filenumber_, _varname_

<br/>

The **Line Input #** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
| _filenumber_|Required. Any valid [file number](../../Glossary/vbe-glossary.md#file-number).|
| _varname_|Required. Valid [Variant](../../Glossary/vbe-glossary.md#variant-data-type) or **String** variable name.|

## Remarks

Data read with **Line Input #** is usually written from a file with **Print #**.

The **Line Input #** statement reads from a file one character at a time until it encounters a carriage return (**Chr**(13)) or carriage return-linefeed (**Chr**(13) + **Chr**(10)) sequence. Carriage return-linefeed sequences are skipped rather than appended to the character string.

## Example

This example uses the **Line Input #** statement to read a line from a sequential file and assign it to a variable. This example assumes that `TESTFILE` is a text file with a few lines of sample data.


```vb
Dim TextLine 
Open "TESTFILE" For Input As #1 ' Open file. 
Do While Not EOF(1) ' Loop until end of file. 
 Line Input #1, TextLine ' Read line into variable. 
 Debug.Print TextLine ' Print to the Immediate window. 
Loop 
Close #1 ' Close file. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
