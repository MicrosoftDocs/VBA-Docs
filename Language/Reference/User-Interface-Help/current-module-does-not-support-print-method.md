---
title: Current module does not support Print method
keywords: vblr6.chm1011234
f1_keywords:
- vblr6.chm1011234
ms.prod: office
ms.assetid: 30f14bb8-ebc6-cbd7-e1f2-e557836c93a9
ms.date: 06/08/2017
localization_priority: Normal
---


# Current module does not support Print method

Not all [methods](../../Glossary/vbe-glossary.md#method) and [properties](../../Glossary/vbe-glossary.md#property) are appropriate in all[modules](../../Glossary/vbe-glossary.md#module). This error has the following causes and solutions:



- You tried to use the **Print** method on an object that can't display anything. For example, you can't use the **Print** method without qualification in a [standard module](../../Glossary/vbe-glossary.md#standard-module).
    
    Remove the reference to the **Print** method, or qualify it with an appropriate object. For example, qualify it with the **Debug** object to display its arguments in the Immediate window during debugging.
    
- You tried to use the **Line**, **Circle**, **PSet**, or **Scale** method on an object that can't accept them. For example, they can't appear unqualified in a standard module or an Automation[class module](../../Glossary/vbe-glossary.md#class-module).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]