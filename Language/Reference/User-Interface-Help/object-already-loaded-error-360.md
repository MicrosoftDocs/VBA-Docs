---
title: Object already loaded (Error 360)
keywords: vblr6.chm1117810
f1_keywords:
- vblr6.chm1117810
ms.prod: office
ms.assetid: e492bbbc-572d-af2f-111f-1879c7b35ea3
ms.date: 06/08/2017
localization_priority: Normal
---


# Object already loaded (Error 360)

The control in the [control array](../../Glossary/vbe-glossary.md#control-array) has already been loaded. This error has the following cause and solution:



- You tried to add a control to a control array at [run time](../../Glossary/vbe-glossary.md#run-time) with the **Load** statement but the index value you referred to already exists.
    
    Change the index reference to a new value or check whether your code is executing the same  **Load** statement with the same index value reference more than once.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]