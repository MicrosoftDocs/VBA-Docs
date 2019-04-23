---
title: ByRef argument type mismatch
keywords: vblr6.chm1011308
f1_keywords:
- vblr6.chm1011308
ms.prod: office
ms.assetid: 6adca657-8620-e3f1-3587-e317f988979c
ms.date: 06/08/2017
localization_priority: Priority
---


# ByRef argument type mismatch

An [argument](../../Glossary/vbe-glossary.md#argument) passed **ByRef** ([by reference](../../Glossary/vbe-glossary.md#by-reference)), the default, must have the precise [data type](../../Glossary/vbe-glossary.md#data-type) expected in the [procedure](../../Glossary/vbe-glossary.md#procedure). This error has the following cause and solution:

- You passed an argument of one type that could not be coerced to the type expected. 
    
  For example, this error occurs if you try to pass an **Integer** variable when a **Long** is expected. If you want coercion to occur, even if it causes information to be lost, you can pass the argument in its own set of parentheses. 
  
  For example, to pass the **Variant** argument `MyVar` to a procedure that expects an **Integer** argument, you can write the call as follows:
    
  ```vb
    Dim MyVar 
    MyVar = 3.1415 
    Call SomeSub((MyVar)) 
    
    Sub SomeSub (MyNum As Integer) 
    MyNum = MyNum + MyNum 
    End Sub
  ```

  Placing the argument in its own set of parentheses forces evaluation of it as an [expression](../../Glossary/vbe-glossary.md#expression). During this evaluation, the fractional portion of the number is rounded (not truncated) to make it conform to the expected argument type. The result of the evaluation is placed in a temporary location, and a reference to the temporary location is received by the procedure. Thus, the original  `MyVar` retains its value.
    
  > [!NOTE] 
  > If you don't specify a type for a [variable](../../Glossary/vbe-glossary.md#variable), the variable receives the default type, **Variant**. This isn't always obvious. For example, the following code declares two variables, the first, `MyVar`, is a **Variant**; the second, `AnotherVar`, is an **Integer**.
  > 
  > ```vb
  >  Dim MyVar, AnotherVar As Integer 
  > ```

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
