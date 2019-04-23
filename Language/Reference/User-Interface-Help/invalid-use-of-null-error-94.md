---
title: Invalid use of Null (Error 94)
keywords: vblr6.chm1000094
f1_keywords:
- vblr6.chm1000094
ms.prod: office
ms.assetid: c1c987fb-8b4c-bbc2-a69b-c5e9047bb94a
ms.date: 06/08/2017
localization_priority: Normal
---


# Invalid use of Null (Error 94)

[Null](../../Glossary/vbe-glossary.md#null) is a **Variant** subtype used to indicate that a data item contains no valid data. This error has the following cause and solution:

- You are trying to obtain the value of a **Variant** variable or an [expression](../../Glossary/vbe-glossary.md#expression) that is **Null**. For example:
    
  ```vb
      MyVar = Null 
    For Count = 1 To MyVar 
    . . . 
    Next Count 
  ```

  Make sure the [variable](../../Glossary/vbe-glossary.md#variable) contains a valid value.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
