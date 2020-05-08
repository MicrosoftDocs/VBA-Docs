---
title: Method or data member not found (Error 461)
keywords: vblr6.chm1000461
f1_keywords:
- vblr6.chm1000461
ms.prod: office
ms.assetid: 10733744-502f-06b3-f0c6-5f039d017be4
ms.date: 06/08/2017
localization_priority: Normal
---


# Method or data member not found (Error 461)

The [collection](../../Glossary/vbe-glossary.md#collection), object, or [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type) doesn't contain the referenced [member](../../Glossary/vbe-glossary.md#member). This error has the following causes and solutions:



- You misspelled the object or member name. Check the spelling of the names and check the **Type** statement or the object documentation to determine what the members are and the proper spelling of the object or member names.
    
- You specified a collection index that's out of range. Check the **Count** property to determine whether a collection member exists. Note that collection indexes begin at 1 rather than zero, so the **Count** property returns the highest possible index number.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
