---
title: Switch function (Visual Basic for Applications)
keywords: vblr6.chm1010096
f1_keywords:
- vblr6.chm1010096
ms.prod: office
ms.assetid: 458ebfcb-af87-1c3b-3f4b-5f308aefa7d9
ms.date: 12/13/2018
localization_priority: Normal
---


# Switch function

Evaluates a list of [expressions](../../Glossary/vbe-glossary.md#expression) and returns a **Variant** value or an expression associated with the first expression in the list that is **True**.

## Syntax

**Switch**(_expr-1_, _value-1_, [ _expr-2_, _value-2_…, [ _expr-n_, _value-n_ ]])

<br/>

The **Switch** function syntax has these parts:

|Part|Description|
|:-----|:-----|
| _expr_|Required. [Variant expression](../../Glossary/vbe-glossary.md#variant-expression) that you want to evaluate.|
| _value_|Required. Value or expression to be returned if the corresponding expression is **True**.|

## Remarks

The **Switch** function [argument](../../Glossary/vbe-glossary.md#argument) list consists of pairs of expressions and values. The expressions are evaluated from left to right, and the value associated with the first expression to evaluate to **True** is returned. 

If the parts aren't properly paired, a [run-time error](../../Glossary/vbe-glossary.md#run-time-error) occurs. For example, if _expr-1_ is **True**, **Switch** returns _value-1_. If _expr-1_ is **False**, but _expr-2_ is **True**, **Switch** returns _value-2_, and so on.

**Switch** returns a [Null](../../Glossary/vbe-glossary.md#null) value if:

- None of the expressions is **True**.
    
- The first **True** expression has a corresponding value that is **Null**.
    
**Switch** evaluates all of the expressions, even though it returns only one of them. For this reason, you should watch for undesirable side effects. For example, if the evaluation of any expression results in a division by zero error, an error occurs.

## Example

This example uses the **Switch** function to return the name of a language that matches the name of a city.


```vb
Function MatchUp(CityName As String)
    Matchup = Switch(CityName = "London", "English", CityName _
                    = "Rome", "Italian", CityName = "Paris", "French")
End Function
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
