---
title: Passing arguments efficiently (VBA)
keywords: vbcn6.chm1009795
f1_keywords:
- vbcn6.chm1009795
ms.prod: office
ms.assetid: cd30d383-05b7-95ed-fd99-5a22c7fe7aab
ms.date: 12/21/2018
localization_priority: Normal
---


# Passing arguments efficiently

All [arguments](../../Glossary/vbe-glossary.md#argument) are passed to [procedures](../../Glossary/vbe-glossary.md#procedure) [by reference](../../Glossary/vbe-glossary.md#by-reference), unless you specify otherwise. This is efficient because all arguments passed by reference take the same amount of time to pass and the same amount of space (4 bytes) within a procedure regardless of the argument's [data type](../../reference/user-interface-help/data-type-summary.md).

You can pass an argument [by value](../../Glossary/vbe-glossary.md#by-value) if you include the **ByVal** keyword in the procedure's declaration. Arguments passed by value consume from 2&ndash;16 bytes within the procedure, depending on the argument's data type. Larger data types take slightly longer to pass by value than smaller ones. Because of this, **String** and **Variant** data types generally should not be passed by value.

Passing an argument by value copies the original [variable](../../Glossary/vbe-glossary.md#variable). Changes to the argument within the procedure aren't reflected back to the original variable. For example:

```vb
Function Factorial (ByVal MyVar As Integer) ' Function declaration. 
 MyVar = MyVar - 1 
 If MyVar = 0 Then 
 Factorial = 1 
 Exit Function 
 End If 
 Factorial = Factorial(MyVar) * (MyVar + 1) 
End Function 
 
' Call Factorial with a variable S. 
S = 5 
Print Factorial(S) ' Displays 120 (the factorial of 5) 
Print S ' Displays 5. 

```

Without including **ByVal** in the function declaration, the preceding **Print** statements would display 1 and 0. This is because `MyVar` would then refer to variable `S`, which is reduced by 1 until it equals 0.

Because **ByVal** makes a copy of the argument, it allows you to pass a variant to the **Factorial** function. You can't pass a variant by reference if the procedure that declares the argument is another data type.

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]