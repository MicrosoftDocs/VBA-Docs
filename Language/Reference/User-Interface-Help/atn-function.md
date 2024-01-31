---
title: Atn function (Visual Basic for Applications)
keywords: vblr6.chm1008860
f1_keywords:
- vblr6.chm1008860
ms.assetid: ab5272cf-b372-8665-28c6-ee0318aa9bac
ms.date: 12/11/2018
ms.localizationpriority: medium
---


# Atn function

Returns a **Double** specifying the arctangent of a number.

## Syntax

**Atn**(_number_)

The required _number_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Double](../../Glossary/vbe-glossary.md#double-data-type) or any valid [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression).

## Remarks

The **Atn** function takes the ratio of two sides of a right triangle (_number_) and returns the corresponding angle in radians. The ratio is the length of the side opposite the angle divided by the length of the side adjacent to the angle.

> [NOTE!] 
> This is the equivalent of the **ATAN()** function in Excel. Oddly, **Excel.WorksheetFunctions.ATAN()** does not exist (as with other Excel formula functions). You can instead use this VBA **Atn()** function.

The range of the result is **-**[pi](../../Glossary/vbe-glossary.md#pi)/2 to pi/2 radians. To convert degrees to radians, multiply degrees by pi/180. To convert radians to degrees, multiply radians by 180/pi.

> [!NOTE] 
> **Atn** is the inverse trigonometric function of **[Tan](tan-function.md)**, which takes an angle as its argument and returns the ratio of two sides of a right triangle. Do not confuse **Atn** with the cotangent, which is the simple inverse of a tangent (1/tangent).

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
