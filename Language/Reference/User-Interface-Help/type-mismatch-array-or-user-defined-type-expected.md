---
title: "Type mismatch: array or user-defined type expected"
keywords: vblr6.chm1011306
f1_keywords:
- vblr6.chm1011306
ms.prod: office
ms.assetid: 31786025-b2c7-8046-4c15-f6bdfad54778
ms.date: 06/08/2017
---


# Type mismatch: array or user-defined type expected

The type of an [argument](../../Glossary/vbe-glossary.md) or[parameter](../../Glossary/vbe-glossary.md) includes whether or not it is an[array](../../Glossary/vbe-glossary.md) or a[user-defined type](../../Glossary/vbe-glossary.md). This error has the following cause and solution:



- Your argument specified a single element of an array or user-defined type, or a simple [variable](../../Glossary/vbe-glossary.md), literal, or [constant](../../Glossary/vbe-glossary.md). However, it is being passed to a parameter that expects a whole array or user-defined type.
    
    Either change the argument or change the definition of the parameter.
    
- Your argument specified an array or user-defined type, but it was not of the same type as the parameter. Either pass an array of the expected type or change the definition of the parameter declaration.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

