---
title: Qualifier must be collection
keywords: vblr6.chm1040118
f1_keywords:
- vblr6.chm1040118
ms.prod: office
ms.assetid: 70c3ce6f-13ca-d9cd-d60c-26c19f803fd7
ms.date: 06/08/2017
localization_priority: Normal
---


# Qualifier must be collection

The use of an exclamation point between two [identifiers](../../Glossary/vbe-glossary.md#identifier) is specific to[collections](../../Glossary/vbe-glossary.md#collection). This error has the following cause and solution:



- You used a name on the left side of the exclamation point (**!**) that isn't the name of a collection. If the name is supposed to represent a collection, check to make sure the name is spelled correctly. Note that the exclamation point is also the [type-declaration character](../../Glossary/vbe-glossary.md#type-declaration-character) for the **Single** data type. If the name in question isn't supposed to be a collection, perhaps the **!** type-declaration character appended to a [variable](../../Glossary/vbe-glossary.md#variable) name has been concatenated with another name.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]