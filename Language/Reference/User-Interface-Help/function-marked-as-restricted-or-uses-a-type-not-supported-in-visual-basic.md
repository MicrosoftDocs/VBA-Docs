---
title: Function marked as restricted or uses a type not supported in Visual Basic
keywords: vblr6.chm1035026
f1_keywords:
- vblr6.chm1035026
ms.prod: office
ms.assetid: b013d6ca-2e99-f2c9-d64b-87ef0990493d
ms.date: 06/08/2017
localization_priority: Normal
---


# Function marked as restricted or uses a type not supported in Visual Basic

Not every [procedure](../../Glossary/vbe-glossary.md#procedure) that appears in a [type library](../../Glossary/vbe-glossary.md#type-library) or [object library](../../Glossary/vbe-glossary.md#object-library) can be accessed by every programming language. The creator of a type or object library can designate some functions as restricted to prevent their use by macro languages. This error has the following causes and solutions:



- You tried to use a function with a restricted specification. You can't use the function in your program. If you have documentation for the object represented by the library, check to see if a [method](../../Glossary/vbe-glossary.md#method) is provided that gives equivalent functionality.
    
- You tried to use a function that requires a [parameter](../../Glossary/vbe-glossary.md#parameter) type or has a return type that isn't available in Visual Basic.
    
    Sometimes you can simulate return types with Visual Basic equivalents. Check the subtypes of the [Variant data type](../../Glossary/vbe-glossary.md#variant-data-type) . This may also work for non-Basic parameter types that are expected as references. However, you can't pass a **Variant** data type[by value](../../Glossary/vbe-glossary.md#by-value) in an effort to simulate a non-Basic type.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]