---
title: Object doesn't support this property or method (Error 438)
keywords: vblr6.chm1011328
f1_keywords:
- vblr6.chm1011328
ms.prod: office
ms.assetid: 0fbab746-dc6d-b227-429a-1f56bb4ca448
ms.date: 06/08/2017
localization_priority: Normal
---


# Object doesn't support this property or method (Error 438)

Not all objects support all [properties](../../Glossary/vbe-glossary.md#property) and [methods](../../Glossary/vbe-glossary.md#method). This error has the following cause and solution:



- You specified a method or property that doesn't exist for this [Automation object](../../Glossary/vbe-glossary.md#automation-object).
    
    See the object's documentation for more information on the object and check the spellings of properties and methods.
    
- You specified a **Friend** procedure to be called late bound. The name of a **Friend** procedure must be known at [compile time](../../Glossary/vbe-glossary.md#compile-time). It can't appear in a late-bound call.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
