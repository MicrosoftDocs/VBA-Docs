---
title: Tan function (Visual Basic for Applications)
keywords: vblr6.chm1009040
f1_keywords:
- vblr6.chm1009040
ms.prod: office
ms.assetid: 4f567334-c397-ccd3-48c9-c42cc630cc79
ms.date: 12/13/2018
localization_priority: Normal
---


# Tan function

Returns a **Double** specifying the tangent of an angle.

## Syntax

**Tan**(_number_)

The required _number_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Double](../../Glossary/vbe-glossary.md#double-data-type) or any valid [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) that expresses an angle in radians.

## Remarks

**Tan** takes an angle and returns the ratio of two sides of a right triangle. The ratio is the length of the side opposite the angle divided by the length of the side adjacent to the angle.

To convert degrees to radians, multiply degrees by [pi](../../Glossary/vbe-glossary.md#pi)/180. To convert radians to degrees, multiply radians by 180/pi.

## Example

This example uses the **Tan** function to return the tangent of an angle.

```vb
Dim MyAngle, MyCotangent
MyAngle = 1.3    ' Define angle in radians.
MyCotangent = 1 / Tan(MyAngle)    ' Calculate cotangent.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]