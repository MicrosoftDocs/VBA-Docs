---
title: Type-declaration character required
keywords: vblr6.chm1011288
f1_keywords:
- vblr6.chm1011288
ms.prod: office
ms.assetid: d4da8cd8-bb2f-d536-7d4a-b7bf701bd361
ms.date: 06/08/2017
localization_priority: Normal
---


# Type-declaration character required

The necessity of using [type-declaration characters](../../Glossary/vbe-glossary.md#type-declaration-character) depends on the form of the [identifier's](../../Glossary/vbe-glossary.md#identifier) declaration. This error has the following cause and solution:

- A [variable](../../Glossary/vbe-glossary.md#variable) that was originally implicitly declared with a type-declaration characters was referenced without a type-declaration character. For example:
    
  ```vb
      MyStr$ = "Implicit declaration" 
    MyStr = "Trying to refer to MyStr$, but error results" _ 
    & "from calling it MyStr." 
  ```

  ```vb
    Dim MyStr$  
    MyStr = "Because it was explicitly declared, the $ is optional." 
  ```

  Either make the declaration explicit, or add the type-declaration character to later references.
    
  > [!NOTE] 
  > If an explicit variable declaration contains a type-declaration character, inclusion of the character is optional in later references.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]