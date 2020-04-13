---
title: Too many local, nonstatic variables
keywords: vblr6.chm1018965
f1_keywords:
- vblr6.chm1018965
ms.prod: office
ms.assetid: 009374ba-1cf5-e4dc-f487-1865bf79de2e
ms.date: 06/08/2017
localization_priority: Normal
---


# Too many local, nonstatic variables

Local, nonstatic [variables](../../Glossary/vbe-glossary.md#variable) are variables that are defined within a [procedure](../../Glossary/vbe-glossary.md#procedure) and reinitialized each time the procedure is called. This error has the following cause and solution:



- The sum of the memory requirements for this procedure's local, nonstatic variables and compiler-generated temporary variables exceeds 32K. Declare some of your variables with the **Static** statement where appropriate. **Static** variables retain their value between procedure invocations because they are allocated from different memory resources than nonstatic variables.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]