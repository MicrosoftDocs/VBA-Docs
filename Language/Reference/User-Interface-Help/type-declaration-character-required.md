---
title: Type-declaration character required
keywords: vblr6.chm1011288
f1_keywords:
- vblr6.chm1011288
ms.prod: office
ms.assetid: d4da8cd8-bb2f-d536-7d4a-b7bf701bd361
ms.date: 06/08/2017
---


# Type-declaration character required

<<<<<<< HEAD
The necessity of using [type-declaration characters](../../Glossary/vbe-glossary.md) depends on the form of the [identifier's](../../Glossary/vbe-glossary.md) declaration. This error has the following cause and solution:

- A [variable](../../Glossary/vbe-glossary.md) that was originally implicitly declared with a type-declaration characters was referenced without a type-declaration character. For example:
=======
The necessity of using [type-declaration characters](../../Glossary/vbe-glossary.md#type-declaration-character) depends on the form of the [identifier's](../../Glossary/vbe-glossary.md#identifier) declaration. This error has the following cause and solution:

- A [variable](../../Glossary/vbe-glossary.md#variable) that was originally implicitly declared with a type-declaration characters was referenced without a type-declaration character. For example:
>>>>>>> master
    
  ```vb
      MyStr$ = "Implicit declaration" 
    MyStr = "Trying to refer to MyStr$, but error results" _ 
<<<<<<< HEAD
    &; "from calling it MyStr." 
=======
    & "from calling it MyStr." 
>>>>>>> master
  ```

  ```vb
    Dim MyStr$  
    MyStr = "Because it was explicitly declared, the $ is optional." 
  ```

  Either make the declaration explicit, or add the type-declaration character to later references.
    
  > [!NOTE] 
  > If an explicit variable declaration contains a type-declaration character, inclusion of the character is optional in later references.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

