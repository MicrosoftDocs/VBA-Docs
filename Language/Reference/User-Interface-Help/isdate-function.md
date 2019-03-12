---
title: IsDate function (Visual Basic for Applications)
keywords: vblr6.chm1008951
f1_keywords:
- vblr6.chm1008951
ms.prod: office
ms.assetid: 832486a7-c69f-8d3b-f0fc-2f6a2f707ecc
ms.date: 12/13/2018
localization_priority: Normal
---


# IsDate function

Returns **True** if the expression is a date or is recognizable as a valid date or time; otherwise, it returns **False**.

## Syntax

**IsDate**(_expression_)

The required _expression_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) containing a [date expression](../../Glossary/vbe-glossary.md#date-expression) or [string expression](../../Glossary/vbe-glossary.md#string-expression) recognizable as a date or time.

## Remarks

In Windows, the range of valid dates is January 1, 100 A.D., through December 31, 9999 A.D.; the ranges vary among operating systems.

## Example

This example uses the **IsDate** function to determine if an expression is recognized as a date or time value.


```vb
Dim MyVar, MyCheck
MyVar = "04/28/2014"    ' Assign valid date value.
MyCheck = IsDate(MyVar)    ' Returns True.

MyVar = "April 28, 2014"    ' Assign valid date value.
MyCheck = IsDate(MyVar)    ' Returns True.

MyVar = "13/32/2014"    ' Assign invalid date value.
MyCheck = IsDate(MyVar)    ' Returns False.

MyVar = "04.28.14"    ' Assign valid time value.
MyCheck = IsDate(MyVar)    ' Returns True.

MyVar = "04.28.2014"    ' Assign invalid time value.
MyCheck = IsDate(MyVar)    ' Returns False.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
