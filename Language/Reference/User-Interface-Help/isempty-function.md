---
title: IsEmpty Function
keywords: vblr6.chm1008952
f1_keywords:
- vblr6.chm1008952
ms.prod: office
ms.assetid: 3fcfe5c2-cc97-17b9-28ca-a47d871a5f1a
ms.date: 06/08/2017
---


# IsEmpty Function



Returns a  **Boolean** value indicating whether a[variable](../../Glossary/vbe-glossary.md#variable) has been initialized.

## Syntax

**IsEmpty(**_expression_**)**
The required  _expression_[argument](../../Glossary/vbe-glossary.md#argument) is a[Variant](../../Glossary/vbe-glossary.md) containing a[numeric](../../Glossary/vbe-glossary.md) or[string expression](../../Glossary/vbe-glossary.md#string-expression). However, because  **IsEmpty** is used to determine if individual variables are initialized, the _expression_ argument is most often a single variable name.

## Remarks

**IsEmpty** returns **True** if the variable is uninitialized, or is explicitly set to[Empty](../../Glossary/vbe-glossary.md#empty); otherwise, it returns  **False**. **False** is always returned if _expression_ contains more than one variable. **IsEmpty** only returns meaningful information for[variants](../../Glossary/vbe-glossary.md).

## Example

This example uses the  **IsEmpty** function to determine whether a variable has been initialized.


```vb
Dim MyVar, MyCheck
MyCheck = IsEmpty(MyVar)    ' Returns True.

MyVar = Null    ' Assign Null.
MyCheck = IsEmpty(MyVar)    ' Returns False.

MyVar = Empty    ' Assign Empty.
MyCheck = IsEmpty(MyVar)    ' Returns True.


```


