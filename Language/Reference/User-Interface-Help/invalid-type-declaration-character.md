---
title: Invalid type-declaration character
keywords: vblr6.chm1011193
f1_keywords:
- vblr6.chm1011193
ms.prod: office
ms.assetid: 6c6411c0-6ed1-3cdb-061b-563ed3b91766
ms.date: 06/08/2017
localization_priority: Normal
---


# Invalid type-declaration character

[Type-declaration characters](../../Glossary/vbe-glossary.md#type-declaration-character) are valid, but don't exist for all[data types](../../Glossary/vbe-glossary.md#data-type); they aren't permitted in some situations. This error has the following causes and solutions:



- A type-declaration character is appended to a [variable](../../Glossary/vbe-glossary.md#variable) declared in a **Private**, **Public**, or **Static** statement with an **As** clause.
    
    Remove the type-declaration character.
    
- A type-declaration character is appended to an inconsistent literal. For example, since the ampersand (**&**) is the type-declaration character for a **Long** integer, appending it to a literal of a different type causes this error:
    
  ```vb
  10.253& 

  ```


     Remove the type-declaration character or replace it with the correct one.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]