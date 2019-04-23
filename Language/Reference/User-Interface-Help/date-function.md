---
title: Date function (Visual Basic for Applications)
keywords: vblr6.chm1008886
f1_keywords:
- vblr6.chm1008886
ms.prod: office
ms.assetid: 8afd02c8-c5b5-f8f3-ff8e-9a2ac0ea94b9
ms.date: 12/12/2018
localization_priority: Normal
---


# Date function

Returns a **Variant** (**Date**) containing the current system date.

## Syntax

**Date**

## Remarks

To set the system date, use the **Date** statement.

**Date**, and if the calendar is Gregorian, **Date$** behavior is unchanged by the **Calendar** property setting. If the calendar is Hijri, **Date$** returns a 10-character string of the form _mm-dd-yyyy_, where _mm_ (01&ndash;12), _dd_ (01&ndash;30) and _yyyy_ (1400&ndash;1523) are the Hijri month, day, and year. The equivalent Gregorian range is Jan 1, 1980, through Dec 31, 2099.

## Example

This example uses the **Date** function to return the current system date.

```vb
Dim MyDate
MyDate = Date    ' MyDate contains the current system date.

```


## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
