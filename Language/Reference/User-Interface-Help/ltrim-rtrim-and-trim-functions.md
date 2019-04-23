---
title: LTrim, RTrim, and Trim functions (Visual Basic for Applications)
keywords: vblr6.chm1008797
f1_keywords:
- vblr6.chm1008797
ms.prod: office
ms.assetid: ffe13d6f-8e7a-3413-98a1-3263c771178b
ms.date: 12/13/2018
localization_priority: Normal
---


# LTrim, RTrim, and Trim functions

Returns a **Variant** (**String**) containing a copy of a specified string without leading spaces (**LTrim**), trailing spaces (**RTrim**), or both leading and trailing spaces (**Trim**).

## Syntax

**LTrim**(_string_)<br/>
**RTrim**(_string_)<br/>
**Trim**(_string_)

The required _string_ [argument](../../Glossary/vbe-glossary.md#argument) is any valid [string expression](../../Glossary/vbe-glossary.md#string-expression). If _string_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned.

## Example

This example uses the **LTrim** function to strip leading spaces, and the **RTrim** function to strip trailing spaces from a string variable. It uses the **Trim** function to strip both types of spaces.


```vb
Dim MyString, TrimString
MyString = "  <-Trim->  "    ' Initialize string.
TrimString = LTrim(MyString)    ' TrimString = "<-Trim->  ".
TrimString = RTrim(MyString)    ' TrimString = "  <-Trim->".
TrimString = LTrim(RTrim(MyString))    ' TrimString = "<-Trim->".
' Using the Trim function alone achieves the same result.
TrimString = Trim(MyString)    ' TrimString = "<-Trim->".


```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
