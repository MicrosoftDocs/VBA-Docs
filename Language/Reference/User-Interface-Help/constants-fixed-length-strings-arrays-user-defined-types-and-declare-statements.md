---
title: Constants, fixed-length strings, arrays, user-defined types, and Declare statements not allowed as Public members of an object module
keywords: vblr6.chm1015643
f1_keywords:
- vblr6.chm1015643
ms.prod: office
ms.assetid: e89fa990-3b88-da3c-961d-a5554eb78ee5
ms.date: 06/08/2017
localization_priority: Normal
---


# Constants, fixed-length strings, arrays, user-defined types, and Declare statements not allowed as Public members of an object module

Not all [variables](../../Glossary/vbe-glossary.md#variable) in an [object module](../../Glossary/vbe-glossary.md#object-module) can be declared as **Public**. However, procedures are **Public** by default, and **Property** procedures can be used to simulate variables syntactically. This error has the following causes and solutions:



- You declared a  **Public**[constant](../../Glossary/vbe-glossary.md#constant) in an object module.
    
    Although you can't declare a  **Public** constant in an object module, you can create a **Property Get** procedure with the same name. If you don't create a **Property Let** or **Property Set** procedure with that name, you are in effect creating a read-only property that can be used the same way you would use a constant.
    
- You declared a  **Public** fixed-length string in an object module. You can simulate fixed-length strings with a set of **Property** procedures that either truncate the string data when it exceeds the permitted length, or notify the user that the length has been exceeded.
    
- You declared a  **Public**[array](../../Glossary/vbe-glossary.md#array) in an object module.
    
    Although a procedure can't return an array, it can return a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) that contains an array. To simulate a **Public** array in a class module, use a set of **Property** procedures that accept and return a **Variant** containing an array.
    
- You placed a  **Declare** statement in an object module. **Declare** statements are implicitly public. Precede the **Declare** statement with the **Private**[keyword](../../Glossary/vbe-glossary.md#keyword).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]