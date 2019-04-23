---
title: Time statement (VBA)
keywords: vblr6.chm1009042
f1_keywords:
- vblr6.chm1009042
ms.prod: office
ms.assetid: 9c11edf2-5eac-207a-985e-1e990f3e1b12
ms.date: 12/03/2018
localization_priority: Normal
---


# Time statement

Sets the system time.

## Syntax

**Time** = _time_

The required _time_ [argument](../../Glossary/vbe-glossary.md#argument) is any [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression), [string expression](../../Glossary/vbe-glossary.md#string-expression), or any combination, that can represent a time.

## Remarks

If _time_ is a string, **Time** attempts to convert it to a time by using the time separators that you specified for your system. If it can't be converted to a valid time, an error occurs.

## Example

This example uses the **Time** statement to set the computer system time to a user-defined time.

```vb
Dim MyTime 
MyTime = #4:35:17 PM# ' Assign a time. 
Time= MyTime ' Set system time to MyTime. 

```

## See also

- [Time function](time-function.md)
- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]