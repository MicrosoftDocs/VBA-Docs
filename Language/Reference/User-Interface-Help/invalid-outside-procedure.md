---
title: Invalid outside procedure
keywords: vblr6.chm1040051
f1_keywords:
- vblr6.chm1040051
ms.prod: office
ms.assetid: 46c00b2b-c656-9ad4-bff9-d341a6a7ecd5
ms.date: 06/08/2017
localization_priority: Normal
---


# Invalid outside procedure

The statement must occur within a  **Sub** or **Function**, or a property procedure (**Property Get**, **Property Let**, **Property Set**). This error has the following cause and solution:



- An executable statement,  **Static** or **ReDim**, appears at [module level](../../Glossary/vbe-glossary.md#module-level).
    
     **Static** is unnecessary at module level, since all module-level[variables](../../Glossary/vbe-glossary.md#variable) are static. Use **Dim** instead of **ReDim** at module level. To create a dynamic[array](../../Glossary/vbe-glossary.md#array) at module level, declare it with **Dim** using empty parentheses.
    
     **Note**  At module level, you can use only [comments](../../Glossary/vbe-glossary.md#comment) and declarative statements, such as **Const**, **Declare**, **Def**_type_, **Dim**, **Option Base**, **Option Compare**, **Option Explicit**, **Option Private**, **Private**, **Public**, and **Type**. The **Sub**, **Function**, and **Property** statements occur outside the body of their[procedures](../../Glossary/vbe-glossary.md#procedure), but within the procedure declaration.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
