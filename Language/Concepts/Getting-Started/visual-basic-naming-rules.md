---
title: Visual Basic naming rules (VBA)
keywords: vbcn6.chm1076688
f1_keywords:
- vbcn6.chm1076688
ms.prod: office
ms.assetid: d3e2b5d5-ac45-a1e0-9935-3630fd033a7d
ms.date: 12/26/2018
localization_priority: Normal
---


# Visual Basic naming rules

Use the following rules when you name [procedures](../../Glossary/vbe-glossary.md#procedure), [constants](../../Glossary/vbe-glossary.md#constant), [variables](../../Glossary/vbe-glossary.md#variable), and [arguments](../../Glossary/vbe-glossary.md#argument) in a Visual Basic [module](../../Glossary/vbe-glossary.md#module):

- You must use a letter as the first character.
    
- You can't use a space, period (**.**), exclamation mark (**!**), or the characters **@**, **&**, **$**, **#** in the name.
    
- Name can't exceed 255 characters in length.
    
- Generally, you shouldn't use any names that are the same as the function, statement, method, and [intrinsic constant](../../Glossary/vbe-glossary.md#intrinsic-constants) names used in Visual Basic or by the [host application](../../Glossary/vbe-glossary.md#host-application). Otherwise you end up shadowing the same [keywords](../../Glossary/vbe-glossary.md#keyword) in the language. To use an intrinsic language function, statement, or method that conflicts with an assigned name, you must explicitly identify it. Precede the intrinsic function, statement, or method name with the name of the associated [type library](../../Glossary/vbe-glossary.md#type-library). For example, if you have a variable called `Left`, you can only invoke the **Left** function by using `VBA.Left`.
    
- You can't repeat names within the same level of [scope](../../Glossary/vbe-glossary.md#scope). For example, you can't declare two variables named `age` within the same procedure. However, you can declare a private variable named `age` and a [procedure-level](../../Glossary/vbe-glossary.md#procedure-level) variable named `age` within the same module.
    
> [!NOTE] 
> Visual Basic isn't case-sensitive, but it preserves the capitalization in the statement where the name is declared.


## See also

- [Document conventions (VBA)](document-conventions-visual-basic-for-applications.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
