---
title: Constant expression required
keywords: vblr6.chm1040114
f1_keywords:
- vblr6.chm1040114
ms.prod: office
ms.assetid: e0493fe4-8f50-c935-391f-0ffaca726b2b
ms.date: 06/08/2017
localization_priority: Normal
---


# Constant expression required

A [constant](../../Glossary/vbe-glossary.md#constant) must be initialized. This error has the following causes and solutions:



- You tried to initialize a constant with a [variable](../../Glossary/vbe-glossary.md#variable), an instance of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type), an object, or the return value of a function call.
    
    Initialize constants with literals, previously declared constants, or literals and constants joined by operators (except the **Is** logical operator).
    
- [array](../../Glossary/vbe-glossary.md#array)
    
    To declare a dynamic array within a [procedure](../../Glossary/vbe-glossary.md#procedure), declare the array with **ReDim** and specify the number of elements with a variable.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
