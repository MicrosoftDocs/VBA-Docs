---
title: Array argument must be ByRef
keywords: vblr6.chm1011079
f1_keywords:
- vblr6.chm1011079
ms.prod: office
ms.assetid: 30259938-07f7-0c89-ccfb-9b16c541e53c
ms.date: 06/08/2017
localization_priority: Normal
---


# Array argument must be ByRef

[Arrays](../../Glossary/vbe-glossary.md#array) declared with **Dim**, **ReDim**, or **Static** can't be passed **ByVal**. This error has the following cause and solution:




- You tried to pass a whole array **ByVal**. An individual element of an array can be passed **ByVal** ([by value](../../Glossary/vbe-glossary.md#by-value)), but a whole array must be passed **ByRef** ([by reference](../../Glossary/vbe-glossary.md#by-reference)). Note that **ByRef** is the default. If you must pass an array **ByVal** to prevent changes to the array's elements from being propagated back to the caller, you can pass the array ([argument](../../Glossary/vbe-glossary.md#argument)) in its own set of parentheses, or you can place it into a **Variant**, and then pass the **Variant** to the **ByVal** parameter, as follows:

    
```vb
Dim MyVar As Variant 
MyVar = OldArray() 

  ```


    
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
