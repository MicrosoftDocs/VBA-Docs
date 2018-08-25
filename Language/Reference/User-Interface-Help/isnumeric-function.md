---
title: IsNumeric Function
keywords: vblr6.chm1008954
f1_keywords:
- vblr6.chm1008954
ms.prod: office
ms.assetid: b8184a41-8400-1228-c40f-1414eb4b6e63
ms.date: 06/08/2017
---


# IsNumeric Function



Returns a  **Boolean** value indicating whether an [expression](../../Glossary/vbe-glossary.md#expression) can be evaluated as a number.

 ## Syntax
 
 **IsNumeric(**_expression_**)**
 
The required  _expression_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Variant](../../Glossary/vbe-glossary.md) containing a [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) or [string expression](../../Glossary/vbe-glossary.md#string-expression).

## Remarks

**IsNumeric** returns **True** if the entire _expression_ is recognized as a number; otherwise, it returns **False**.
 **IsNumeric** returns **False** if _expression_ is a [date expression](../../Glossary/vbe-glossary.md#date-expression).

## Example

This example uses the  **IsNumeric** function to determine if a variable can be evaluated as a number.


```vb
Dim MyVar, MyCheck
MyVar = "53"    ' Assign value.
MyCheck = IsNumeric(MyVar)    ' Returns True.

MyVar = "459.95"    ' Assign value.
MyCheck = IsNumeric(MyVar)    ' Returns True.

MyVar = "45 Help"    ' Assign value.
MyCheck = IsNumeric(MyVar)    ' Returns False.


```


