---
title: UCase function (Visual Basic for Applications)
keywords: vblr6.chm1009051
f1_keywords:
- vblr6.chm1009051
ms.prod: office
ms.assetid: 444bd68b-a2bf-11b2-e6b7-76edf9b03ecd
ms.date: 12/13/2018
localization_priority: Normal
---


# UCase function

Returns a **Variant** (**String**) containing the specified string, converted to uppercase.

## Syntax

**UCase**(_string_)

The required _string_ [argument](../../Glossary/vbe-glossary.md#argument) is any valid [string expression](../../Glossary/vbe-glossary.md#string-expression). If _string_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned.

## Remarks

Only lowercase letters are converted to uppercase; all uppercase letters and nonletter characters remain unchanged.

## Example

This example uses the **UCase** function to return an uppercase version of a string.

```vb
Dim LowerCase, UpperCase
LowerCase = "Hello World 1234"    ' String to convert.
UpperCase = UCase(LowerCase)    ' Returns "HELLO WORLD 1234".

```

## See also

- [LCase function](lcase-function.md)
- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
