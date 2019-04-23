---
title: Space function (Visual Basic for Applications)
keywords: vblr6.chm1009026
f1_keywords:
- vblr6.chm1009026
ms.prod: office
ms.assetid: fa531cfb-863f-ede9-34b8-6000711d71ed
ms.date: 12/13/2018
localization_priority: Normal
---


# Space function

Returns a **Variant** (**String**) consisting of the specified number of spaces.

## Syntax

**Space**(_number_)

The required _number_ [argument](../../Glossary/vbe-glossary.md#argument) is the number of spaces you want in the string.

## Remarks

The **Space** function is useful for formatting output and clearing data in fixed-length strings.

## Example

This example uses the **Space** function to return a string consisting of a specified number of spaces.


```vb
Dim MyString
' Returns a string with 10 spaces.
MyString = Space(10)

' Insert 10 spaces between two strings.
MyString = "Hello" & Space(10) & "World"

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
