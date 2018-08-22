---
title: Invalid use of Me keyword
keywords: vblr6.chm1015641
f1_keywords:
- vblr6.chm1015641
ms.prod: office
ms.assetid: c1751bda-c3f5-84c3-0fe0-4ddcdd4829c6
ms.date: 06/08/2017
---


# Invalid use of Me keyword

The **Me** [keyword](../../Glossary/vbe-glossary.md#keyword) can appear in [class modules](../../Glossary/vbe-glossary.md#class-module). This error has the following causes and solutions:

- The **Me** keyword appeared in a [standard module](../../Glossary/vbe-glossary.md#standard-module).
    
  The **Me** keyword can't appear in a standard module because a standard module doesn't represent an object. If you copied the code in question from a class module, you have to replace the **Me** keyword with the specific object or form name to preserve the original reference.
    
- The **Me** keyword appeared on the left side of a **Set** assignment, for example:
    
  ```vb
    Set Me = MyObject    ' Causes "Invalid use of Me keyword" message. 

  ```

  Remove the **Set** assignment.
    
  > [!NOTE] 
  > The **Me** keyword can appear on the left side of a **Let** assignment, in which case the default [property](../../Glossary/vbe-glossary.md#property) of the object represented by **Me** is set. For example:

  ```vb
    Let Me = MyObject   ' Valid assignment with explicit Let. 
    Me = MyObject       ' Valid assignment with implicit Let. 
  ```

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

