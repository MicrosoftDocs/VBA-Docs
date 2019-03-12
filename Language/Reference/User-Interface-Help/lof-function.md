---
title: LOF function (Visual Basic for Applications)
keywords: vblr6.chm1008965
f1_keywords:
- vblr6.chm1008965
ms.prod: office
ms.assetid: 1bf66bce-d3d7-9c34-e8d2-8ad1e1ee24a8
ms.date: 12/13/2018
localization_priority: Normal
---


# LOF function

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) representing the size, in bytes, of a file opened by using the **[Open](open-statement.md)** statement.

## Syntax

**LOF**(_filenumber_)

The required _filenumber_ [argument](../../Glossary/vbe-glossary.md#argument) is an [Integer](../../Glossary/vbe-glossary.md#integer-data-type) containing a valid [file number](../../Glossary/vbe-glossary.md#file-number).

> [!NOTE] 
> Use the **[FileLen](filelen-function.md)** function to obtain the length of a file that is not open.


## Example

This example uses the **LOF** function to determine the size of an open file. This example assumes that `TESTFILE` is a text file containing sample data.

```vb
Dim FileLength
Open "TESTFILE" For Input As #1    ' Open file.
FileLength = LOF(1)    ' Get length of file.
Close #1    ' Close file.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
