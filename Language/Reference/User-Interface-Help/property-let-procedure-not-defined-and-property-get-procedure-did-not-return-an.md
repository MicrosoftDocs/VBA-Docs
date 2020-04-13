---
title: Property let procedure not defined and property get procedure did not return an object (Error 451)
keywords: vblr6.chm1011233
f1_keywords:
- vblr6.chm1011233
ms.prod: office
ms.assetid: 7f34f9a0-d83a-3fd6-50cd-10f82d893ee1
ms.date: 06/08/2017
localization_priority: Normal
---


# Property let procedure not defined and property get procedure did not return an object (Error 451)

Certain [properties](../../Glossary/vbe-glossary.md#property), [methods](../../Glossary/vbe-glossary.md#method), and operations can only apply to **Collection** objects. This error has the following cause and solution:



- You specified an operation or property that is exclusive to [collections](../../Glossary/vbe-glossary.md#collection), but the object isn't a collection.
    
    Check the spelling of the object or property name, or verify that the object is a **Collection** object. Also look at the **Add** method used to add the object to the collection to be sure the syntax is correct and that any[identifiers](../../Glossary/vbe-glossary.md#identifier) were spelled correctly.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]