---
title: IsError function (Visual Basic for Applications)
keywords: vblr6.chm1008824
f1_keywords:
- vblr6.chm1008824
ms.prod: office
ms.assetid: 7eab8dd7-6719-3fc1-fea2-3140cc6a0e5f
ms.date: 12/13/2018
localization_priority: Normal
---


# IsError function

Returns a **Boolean** value indicating whether an [expression](../../Glossary/vbe-glossary.md#expression) is an error value.

## Syntax

**IsError**(_expression_)

The required _expression_ [argument](../../Glossary/vbe-glossary.md#argument) can be any valid expression.

## Remarks

Error values are created by converting real numbers to error values by using the **[CVErr](cverr-function.md)** function. The **IsError** function is used to determine if a [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) represents an error. **IsError** returns **True** if the _expression_ argument indicates an error; otherwise, it returns **False**.

## Example

This example uses the **IsError** function to check if a numeric expression is an error value. The **CVErr** function is used to return an **Error Variant** from a user-defined function. Assume that `UserFunction` is a user-defined function procedure that returns an error value; for example, a return value assigned with the statement `UserFunction = CVErr(32767)`, where 32767 is a user-defined number.


```vb
Dim ReturnVal, MyCheck
ReturnVal = UserFunction()
MyCheck = IsError(ReturnVal)    ' Returns True.
```


## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
