---
title: FileDateTime function (Visual Basic for Applications)
keywords: vblr6.chm1008921
f1_keywords:
- vblr6.chm1008921
ms.prod: office
ms.assetid: d4a54c4c-dc61-cb70-38b4-9c5506cfe789
ms.date: 12/12/2018
localization_priority: Normal
---


# FileDateTime function

Returns a **Variant** (**Date**) that indicates the date and time when a file was created or last modified.

## Syntax

**FileDateTime**(_pathname_)

The required _pathname_ [argument](../../Glossary/vbe-glossary.md#argument) is a [string expression](../../Glossary/vbe-glossary.md#string-expression) that specifies a file name. The _pathname_ may include the directory or folder, and the drive.

## Example

This example uses the **FileDateTime** function to determine the date and time a file was created or last modified. The format of the date and time displayed is based on the locale settings of your system.


```vb
Dim MyStamp
' Assume TESTFILE was last modified on February 12, 1993 at 4:35:47 PM.
' Assume English/U.S. locale settings.
MyStamp = FileDateTime("TESTFILE")    ' Returns "2/12/93 4:35:47 PM".


```


## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
