---
title: For Each can only iterate over a collection object or an array
keywords: vblr6.chm1011172
f1_keywords:
- vblr6.chm1011172
ms.prod: office
ms.assetid: 0defd9d4-4775-c5dd-1212-951016efe997
ms.date: 06/08/2017
localization_priority: Normal
---


# For Each can only iterate over a collection object or an array

The  **For Each** construct can only be used with [collections](../../Glossary/vbe-glossary.md#collection) and [arrays](../../Glossary/vbe-glossary.md#array). This error has the following cause and solution:



- You specified an object that isn't a collection or array as the  _group_ part of the **For Each** syntax. Check the spelling of the item over which you want to iterate to make sure it corresponds to a collection or array in [scope](../../Glossary/vbe-glossary.md#scope) in this part of your code.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]