---
title: Sin Function
keywords: vblr6.chm1009024
f1_keywords:
- vblr6.chm1009024
ms.prod: office
ms.assetid: 94829038-6b78-3dbf-cfe0-558caf343ff1
ms.date: 06/08/2017
---


# Sin Function



Returns a  **Double** specifying the sine of an angle.

## Syntax

**Sin(**_number_**)**
The required  _number_[argument](../../Glossary/vbe-glossary.md#argument) is a[Double](../../Glossary/vbe-glossary.md) or any valid[numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) that expresses an angle in radians.

## Remarks

The  **Sin** function takes an angle and returns the ratio of two sides of a right triangle. The ratio is the length of the side opposite the angle divided by the length of the hypotenuse.
The result lies in the range -1 to 1.
To convert degrees to radians, multiply degrees by [pi](../../Glossary/vbe-glossary.md#pi)/180. To convert radians to degrees, multiply radians by 180/pi.

## Example

This example uses the  **Sin** function to return the sine of an angle.


```vb
Dim MyAngle, MyCosecant
MyAngle = 1.3    ' Define angle in radians.
MyCosecant = 1 / Sin(MyAngle)    ' Calculate cosecant.

```


