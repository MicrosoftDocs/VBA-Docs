---
title: Cannot define a KWID_tkPUBLIC user-defined type within a private object module
keywords: vblr6.chm1040352
f1_keywords:
- vblr6.chm1040352
ms.prod: office
ms.assetid: 594b1460-9990-57c6-9483-003827033d27
ms.date: 06/08/2017
localization_priority: Normal
---


# Cannot define a KWID_tkPUBLIC user-defined type within a private object module

A [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) that appears within an [object module](../../Glossary/vbe-glossary.md#object-module) can't be **Public**. This error has the following cause and solution:



- You tried to define a **Public** user-defined type in an object module. Move the user-defined type definition to a [standard module](../../Glossary/vbe-glossary.md#standard-module), and then declare [variables](../../Glossary/vbe-glossary.md#variable) of the type in the object module or other[modules](../../Glossary/vbe-glossary.md#module), as appropriate. If you only want the type to be available in the module in which it appears, you can place its **Type...End Type** definition in the object module and precede its definition with the **Private**[keyword](../../Glossary/vbe-glossary.md#keyword).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]