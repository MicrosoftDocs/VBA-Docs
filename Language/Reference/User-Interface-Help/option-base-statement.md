---
title: Option Base statement (VBA)
keywords: vblr6.chm1008990
f1_keywords:
- vblr6.chm1008990
ms.prod: office
ms.assetid: 21f45e9e-2cb2-3a45-0484-d23adae77e3e
ms.date: 12/03/2018
localization_priority: Normal
---


# Option Base statement

Used at the [module level](../../Glossary/vbe-glossary.md#module-level) to declare the default lower bound for [array](../../Glossary/vbe-glossary.md#array) subscripts.

## Syntax

**Option Base** { **0** | **1** }

## Remarks

Because the default base is **0**, the **Option Base** statement is never required. If used, the [statement](../../Glossary/vbe-glossary.md#statement) must appear in a [module](../../Glossary/vbe-glossary.md#module) before any [procedures](../../Glossary/vbe-glossary.md#procedure). **Option Base** can appear only once in a module and must precede array [declarations](../../Glossary/vbe-glossary.md#declaration) that include dimensions.

> [!NOTE] 
> The **To** clause in the **Dim**, **Private**, **Public**, **ReDim**, and **Static** statements provides a more flexible way to control the range of an array's subscripts. However, if you don't explicitly set the lower bound with a **To** clause, you can use **Option Base** to change the default lower bound to 1. The base of an array created with the **ParamArray** keyword is zero; **Option Base** does not affect **ParamArray** (or the **[Array](array-function.md)** function, when qualified with the name of its type library, for example **VBA.Array**).

The **Option Base** statement only affects the lower bound of arrays in the module where the statement is located.

## Example

This example uses the **Option Base** statement to override the default base array subscript value of 0. The **[LBound](lbound-function.md)** function returns the smallest available subscript for the indicated dimension of an array. The **Option Base** statement is used at the module level only.


```vb
Option Base 1 ' Set default array subscripts to 1. 
 
Dim Lower 
Dim MyArray(20), TwoDArray(3, 4) ' Declare array variables. 
Dim ZeroArray(0 To 5) ' Override default base subscript. 
' Use LBound function to test lower bounds of arrays. 
Lower = LBound(MyArray) ' Returns 1. 
Lower = LBound(TwoDArray, 2) ' Returns 1. 
Lower = LBound(ZeroArray) ' Returns 0. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
