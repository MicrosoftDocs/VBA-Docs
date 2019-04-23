---
title: Object data type
keywords: vblr6.chm1008829
f1_keywords:
- vblr6.chm1008829
ms.prod: office
ms.assetid: cffe448d-29dd-52aa-4a5c-2155c07b5bf3
ms.date: 11/19/2018
localization_priority: Normal
---


# Object data type

[Object variables](../../Glossary/vbe-glossary.md#object-variable) are stored as 32-bit (4-byte) addresses that refer to objects. Using the **Set** statement, a variable declared as an **Object** can have any object reference assigned to it.


> [!NOTE] 
> Although a variable declared with **Object** type is flexible enough to contain a reference to any object, binding to the object referenced by that variable is always late ([run-time](../../Glossary/vbe-glossary.md#run-time) binding). 
> 
> To force early binding ([compile-time](../../Glossary/vbe-glossary.md#compile-time) binding), assign the object reference to a variable declared with a specific [class](../../Glossary/vbe-glossary.md#class) name.


## See also

- [Data type summary](data-type-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
