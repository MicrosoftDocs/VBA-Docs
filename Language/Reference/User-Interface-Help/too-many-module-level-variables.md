---
title: Too many module-level variables
keywords: vblr6.chm1018966
f1_keywords:
- vblr6.chm1018966
ms.prod: office
ms.assetid: d07aa660-f8d3-908c-c813-9db33e4a8ac3
ms.date: 06/08/2017
localization_priority: Normal
---


# Too many module-level variables

[Module-level](../../Glossary/vbe-glossary.md#module-level) [variables](../../Glossary/vbe-glossary.md#variable) are those declared in the Declarations section of a [module](../../Glossary/vbe-glossary.md#module), before the module's [procedures](../../Glossary/vbe-glossary.md#procedure). This error has the following cause and solution:



- The sum of the memory requirements for all module-level variables in this [module](../../Glossary/vbe-glossary.md#module) exceeds 64K.
    
    This is the storage limit for this module. If appropriate, you can declare some of your variables as  **Public** in another module, or if some module-level variables are used only in one procedure, you can declare them within that procedure. If you declared variables at module level because you want them to retain their value between procedure invocations, you can instead declare them as **Static** within the procedure in which they are referenced.
    
     **Note**  Available space can vary among operating systems.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]