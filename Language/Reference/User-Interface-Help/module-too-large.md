---
title: Module too large
keywords: vblr6.chm1057026
f1_keywords:
- vblr6.chm1057026
ms.prod: office
ms.assetid: b00483e1-d3b2-f532-eaa3-fae61f45c013
ms.date: 06/08/2017
localization_priority: Normal
---


# Module too large

A [module](../../Glossary/vbe-glossary.md#module) contains code within the [project](../../Glossary/vbe-glossary.md#project). This error has the following cause and solution:



- There is too much code in the module.
    
    Create a new module and move some of the [procedures](../../Glossary/vbe-glossary.md#procedure) from this module to the new one. If the current module contains[module-level](../../Glossary/vbe-glossary.md#module-level) declarations of data that must be visible to the procedures in the new module, declare that data as **Public**.
    
    **Note**  [Comments](../../Glossary/vbe-glossary.md#comment) aren't counted as lines of code. Therefore, deleting comments doesn't prevent this error.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]