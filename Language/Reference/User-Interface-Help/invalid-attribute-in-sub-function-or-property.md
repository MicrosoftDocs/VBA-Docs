---
title: Invalid attribute in sub, function, or property
keywords: vblr6.chm1011196
f1_keywords:
- vblr6.chm1011196
ms.prod: office
ms.assetid: 86a5ff38-4f00-060f-5087-453758f27e68
ms.date: 06/08/2017
localization_priority: Normal
---


# Invalid attribute in sub, function, or property

Some attributes are invalid within [procedures](../../Glossary/vbe-glossary.md#procedure). This error has the following cause and solution:

- A **Public** or **Private** attribute appears within the body of a procedure definition. Remove the attribute from the procedure. To give the [variable](../../Glossary/vbe-glossary.md#variable) wider [scope](../../Glossary/vbe-glossary.md#scope), move the declaration to [module level](../../Glossary/vbe-glossary.md#module-level). Variables declared within procedures are always **Private**.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]