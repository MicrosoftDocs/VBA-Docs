---
title: "'Item' property can't be set at run time (Error 382)"
keywords: vblr6.chm382
f1_keywords:
- vblr6.chm382
ms.prod: office
ms.assetid: 20149505-5b45-6c97-228e-839bee802c62
ms.date: 06/08/2017
localization_priority: Normal
---


# 'Item' property can't be set at run time (Error 382)

The [property](../../Glossary/vbe-glossary.md#property) is read-only at [run time](../../Glossary/vbe-glossary.md#run-time). This error has the following cause and solution:



- You tried to set or change a property whose value can only be set at [design time](../../Glossary/vbe-glossary.md#design-time).
    
    Remove the reference to the property from your code or change the reference to only return the value of the property at run time.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]