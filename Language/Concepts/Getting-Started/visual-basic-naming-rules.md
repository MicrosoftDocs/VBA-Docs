---
title: Visual Basic Naming Rules
keywords: vbcn6.chm1076688
f1_keywords:
- vbcn6.chm1076688
ms.prod: office
ms.assetid: d3e2b5d5-ac45-a1e0-9935-3630fd033a7d
ms.date: 06/08/2017
---


# Visual Basic Naming Rules

<<<<<<< HEAD
Use the following rules when you name [procedures](../../Glossary/vbe-glossary.md), [constants](../../Glossary/vbe-glossary.md), [variables](../../Glossary/vbe-glossary.md), and [arguments](../../Glossary/vbe-glossary.md) in a Visual Basic[module](../../Glossary/vbe-glossary.md):
=======
Use the following rules when you name [procedures](../../Glossary/vbe-glossary.md#procedure), [constants](../../Glossary/vbe-glossary.md#constant), [variables](../../Glossary/vbe-glossary.md#variable), and [arguments](../../Glossary/vbe-glossary.md#argument) in a Visual Basic[module](../../Glossary/vbe-glossary.md#module):
>>>>>>> master



- You must use a letter as the first character.
    
- You can't use a space, period (**.**), exclamation mark (**!**), or the characters **@**, **&;**, **$**, **#** in the name.
    
- Name can't exceed 255 characters in length.
    
<<<<<<< HEAD
- Generally, you shouldn't use any names that are the same as the [functions](../../Glossary/vbe-glossary.md), [statements](../../Glossary/vbe-glossary.md), and [methods](../../Glossary/vbe-glossary.md) in Visual Basic. You end up shadowing the same[keywords](../../Glossary/vbe-glossary.md) in the language. To use an intrinsic language function, statement, or method that conflicts with an assigned name, you must explicitly identify it. Precede the intrinsic function, statement, or method name with the name of the associated[type library](../../Glossary/vbe-glossary.md). For example, if you have a variable called  `Left`, you can only invoke the  **Left** function using `VBA.Left`.
    
- You can't repeat names within the same level of [scope](../../Glossary/vbe-glossary.md). For example, you can't declare two variables named  `age` within the same procedure. However, you can declare a private variable named `age` and a[procedure-level](../../Glossary/vbe-glossary.md) variable named `age` within the same module.
=======
- Generally, you shouldn't use any names that are the same as the function, statement, method, and [intrinsic constant](../../Glossary/vbe-glossary.md#intrinsic-constants) names used in Visual Basic or by the [host application](../../Glossary/vbe-glossary.md#host-application). Otherwise you end up shadowing the same [keywords](../../Glossary/vbe-glossary.md#keyword) in the language. To use an intrinsic language function, statement, or method that conflicts with an assigned name, you must explicitly identify it. Precede the intrinsic function, statement, or method name with the name of the associated [type library](../../Glossary/vbe-glossary.md#type-library). For example, if you have a variable called  `Left`, you can only invoke the  **Left** function using `VBA.Left`.
    
- You can't repeat names within the same level of [scope](../../Glossary/vbe-glossary.md#scope). For example, you can't declare two variables named  `age` within the same procedure. However, you can declare a private variable named `age` and a[procedure-level](../../Glossary/vbe-glossary.md#procedure-level) variable named `age` within the same module.
>>>>>>> master
    
     **Note**  Visual Basic isn't case-sensitive, but it preserves the capitalization in the statement where the name is declared.


