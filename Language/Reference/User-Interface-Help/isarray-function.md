---
title: IsArray Function
keywords: vblr6.chm1008823
f1_keywords:
- vblr6.chm1008823
ms.prod: office
ms.assetid: b7926cce-3e55-4074-1a04-99dac608fcb1
ms.date: 06/08/2017
---


# IsArray Function



Returns a  **Boolean** value indicating whether a[variable](../../Glossary/vbe-glossary.md#variable) is an[array](../../Glossary/vbe-glossary.md#array).

## Syntax

**IsArray(**_varname_**)**
The required  _varname_[argument](../../Glossary/vbe-glossary.md#argument) is an[identifier](../../Glossary/vbe-glossary.md#identifier) specifying a variable.

## Remarks

**IsArray** returns **True** if the variable is an array; otherwise, it returns **False**. **IsArray** is especially useful with[variants](../../Glossary/vbe-glossary.md#variant) containing arrays.

## Example

This example uses the  **IsArray** function to check if a variable is an array.


```vb
Dim MyArray(1 To 5) As Integer, YourArray, MyCheck    ' Declare array variables.
YourArray = Array(1, 2, 3)    ' Use Array function.
MyCheck = IsArray(MyArray)    ' Returns True.
MyCheck = IsArray(YourArray)    ' Returns True.


```


