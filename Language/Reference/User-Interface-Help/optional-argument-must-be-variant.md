---
title: Optional argument must be Variant
keywords: vblr6.chm1011239
f1_keywords:
- vblr6.chm1011239
ms.prod: office
ms.assetid: 24c249a4-f0aa-4437-fb57-9f74c748dddf
ms.date: 06/08/2017
localization_priority: Normal
---


# Optional argument must be Variant

Optional [arguments](../../Glossary/vbe-glossary.md#argument) can have any intrinsic[data type](../../Glossary/vbe-glossary.md#data-type), but it must be a type with a single default value. This error has the following cause and solution:



- You tried to specify  **Optional** with a [parameter](../../Glossary/vbe-glossary.md#parameter) that has no default value, for example, an [array](../../Glossary/vbe-glossary.md#array).
    
    Make sure any argument specified as  **Optional** has a default value.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]