---
title: Circular dependencies between modules
keywords: vblr6.chm1011110
f1_keywords:
- vblr6.chm1011110
ms.prod: office
ms.assetid: 89b0ffde-11a5-9d8b-927c-386abf69f6e7
ms.date: 06/08/2017
---


# Circular dependencies between modules

<<<<<<< HEAD
Circular references between [modules](../../Glossary/vbe-glossary.md), [constants](../../Glossary/vbe-glossary.md), and [user-defined types](../../Glossary/vbe-glossary.md) aren't allowed. This error has the following cause and solution:
=======
Circular references between [modules](../../Glossary/vbe-glossary.md#module), [constants](../../Glossary/vbe-glossary.md#constant), and [user-defined types](../../Glossary/vbe-glossary.md#user-defined-type) aren't allowed. This error has the following cause and solution:
>>>>>>> master



- A user-defined type or constant in one module references a user-defined type or constant in a second module, which in turn references another user-defined type or constant in the first module. Remove the dependent references.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

