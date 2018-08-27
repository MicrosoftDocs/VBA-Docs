---
title: LCase Function
keywords: vblr6.chm1011360
f1_keywords:
- vblr6.chm1011360
ms.prod: office
ms.assetid: aeccc222-c9c7-85e9-fa03-8ac99bcfe9dd
ms.date: 06/08/2017
---


# LCase Function



Returns a [String](../../Glossary/vbe-glossary.md#string-data-type) that has been converted to lowercase.

## Syntax

**LCase** ( _string_ )
The required  _string_[argument](../../Glossary/vbe-glossary.md#argument) is any valid[string expression](../../Glossary/vbe-glossary.md#string-expression). If  _string_ contains[Null](../../Glossary/vbe-glossary.md#null), Null is returned.

## Remarks

Only uppercase letters are converted to lowercase; all lowercase letters and nonletter characters remain unchanged.

## Example

This example uses the  **LCase** function to return a lowercase version of a string.


```vb
Dim UpperCase, LowerCase
Uppercase = "Hello World 1234"    ' String to convert.
Lowercase = Lcase(UpperCase)    ' Returns "hello world 1234".


```


