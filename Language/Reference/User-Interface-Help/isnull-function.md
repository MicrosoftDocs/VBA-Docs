---
title: IsNull function (Visual Basic for Applications)
keywords: vblr6.chm1008953
f1_keywords:
- vblr6.chm1008953
ms.prod: office
ms.assetid: 875909ba-289e-aba9-0462-9327efe0bc46
ms.date: 12/13/2018
localization_priority: Normal
---


# IsNull function

Returns a **Boolean** value that indicates whether an [expression](../../Glossary/vbe-glossary.md#expression) contains no valid data ([Null](../../Glossary/vbe-glossary.md#null)).

## Syntax

**IsNull**(_expression_)

The required _expression_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) containing a [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) or a [string expression](../../Glossary/vbe-glossary.md#string-expression).

## Remarks

**IsNull** returns **True** if _expression_ is **Null**; otherwise, **IsNull** returns **False**. If _expression_ consists of more than one [variable](../../Glossary/vbe-glossary.md#variable), **Null** in any constituent variable causes **True** to be returned for the entire expression.

The **Null** value indicates that the **Variant** contains no valid data. **Null** is not the same as [Empty](../../Glossary/vbe-glossary.md#empty), which indicates that a variable has not yet been initialized. It is also not the same as a zero-length string (""), which is sometimes referred to as a null string.

> [!IMPORTANT] 
> Use the **IsNull** function to determine whether an expression contains a **Null** value. Expressions that you might expect to evaluate to **True** under some circumstances, such as `If Var = Null` and `If Var <> Null`, are always **False**. This is because any expression containing a **Null** is itself **Null** and therefore **False**.

## Example

This example uses the **IsNull** function to determine if a variable contains a **Null**.

```vb
Dim MyVar, MyCheck
MyCheck = IsNull(MyVar)    ' Returns False.

MyVar = ""
MyCheck = IsNull(MyVar)    ' Returns False.

MyVar = Null
MyCheck = IsNull(MyVar)    ' Returns True.

```


## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
