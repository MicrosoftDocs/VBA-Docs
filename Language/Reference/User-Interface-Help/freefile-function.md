---
title: FreeFile function (Visual Basic for Applications)
keywords: vblr6.chm1008926
f1_keywords:
- vblr6.chm1008926
ms.prod: office
ms.assetid: b3fda54f-0cbd-788b-e944-d7d7b07a02a1
ms.date: 12/12/2018
localization_priority: Normal
---


# FreeFile function

Returns an [Integer](../../Glossary/vbe-glossary.md#integer-data-type) representing the next [file number](../../Glossary/vbe-glossary.md#file-number) available for use by the **[Open](open-statement.md)** statement.

## Syntax

**FreeFile** [ (_rangenumber_) ]

The optional _rangenumber_ argument is a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) that specifies the range from which the next free file number is to be returned. Specify a 0 (default) to return a file number in the range 1&ndash;255, inclusive. Specify a 1 to return a file number in the range 256&ndash;511.

## Remarks

Use **FreeFile** to supply a file number that is not already in use.

## Example

This example uses the **FreeFile** function to return the next available file number. Five files are opened for output within the loop, and some sample data is written to each.


```vb
Dim MyIndex, FileNumber
For MyIndex = 1 To 5    ' Loop 5 times.
    FileNumber = FreeFile    ' Get unused file
        ' number.
    Open "TEST" & MyIndex For Output As #FileNumber    ' Create file name.
    Write #FileNumber, "This is a sample."    ' Output text.
    Close #FileNumber    ' Close file.
Next MyIndex


```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
