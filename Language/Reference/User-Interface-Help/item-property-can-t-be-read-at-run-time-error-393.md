---
title: "'Item' property can't be read at run time (Error 393)"
keywords: vblr6.chm1117826
f1_keywords:
- vblr6.chm1117826
ms.prod: office
ms.assetid: 80b33869-c3a3-9f3f-57e4-076b81b31a66
ms.date: 06/08/2017
localization_priority: Normal
---


# 'Item' property can't be read at run time (Error 393)

The [property](../../Glossary/vbe-glossary.md#property) is only available at [design time](../../Glossary/vbe-glossary.md#design-time). This error has the following cause and solution:



- You tried to read a property at [run time](../../Glossary/vbe-glossary.md#run-time) that is only accessible at design time.
    
    Change your code and remove the reference to the property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]