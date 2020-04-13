---
title: Private Enum and user-defined types cannot be used as parameters or return types for public procedures, public data members, or fields of public user-defined types
keywords: vblr6.chm1108114
f1_keywords:
- vblr6.chm1108114
ms.prod: office
ms.assetid: bb291092-bc58-fc0c-9a3e-fdaf84886952
ms.date: 06/08/2017
localization_priority: Normal
---


# Private Enum and user-defined types cannot be used as parameters or return types for public procedures, public data members, or fields of public user-defined types

A **Public** procedure is visible to all[modules](../../Glossary/vbe-glossary.md#module) in a [project](../../Glossary/vbe-glossary.md#project), while a **Private** **Enum** type is not visible outside its own module. This error has the following cause and solution:



- Your **Public** procedure is in a **Public** class, but it returns a value or has a [parameter](../../Glossary/vbe-glossary.md#parameter) that is defined in a [standard module](../../Glossary/vbe-glossary.md#standard-module) or in a **Private** class.
    
    Declare the **Enum** **Public**. It must be in a [class module](../../Glossary/vbe-glossary.md#class-module).
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]