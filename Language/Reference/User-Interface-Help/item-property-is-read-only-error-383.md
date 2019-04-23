---
title: "'Item' property is read-only (Error 383)"
keywords: vblr6.chm1117824
f1_keywords:
- vblr6.chm1117824
ms.prod: office
ms.assetid: 6ef3eb14-5e32-5639-e297-990184249393
ms.date: 06/08/2017
localization_priority: Normal
---


# 'Item' property is read-only (Error 383)

The [property](../../Glossary/vbe-glossary.md#property) is read-only at both[design time](../../Glossary/vbe-glossary.md#design-time) and [run time](../../Glossary/vbe-glossary.md#run-time). This error has the following cause and solution:



- You tried to set or change a property whose value can only be read. Remove the reference to the property from your code or change the reference to only return the value of the property at run time.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]