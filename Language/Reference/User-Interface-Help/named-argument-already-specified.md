---
title: Named argument already specified
keywords: vblr6.chm1011224
f1_keywords:
- vblr6.chm1011224
ms.prod: office
ms.assetid: 8fa1e0f1-2484-8344-038c-438ab21d2b71
ms.date: 06/08/2017
localization_priority: Normal
---


# Named argument already specified

You can use a [named argument](../../Glossary/vbe-glossary.md#named-argument) only once in the [argument](../../Glossary/vbe-glossary.md#argument) list of each[procedure](../../Glossary/vbe-glossary.md#procedure) invocation. This error has the following causes and solutions:



- You specified the same named argument more than once in a single call. For example, if the procedure  `MySub` expects the named arguments `Arg1` and `Arg2`, the following call would generate this error:
    
  ```vb
  Call MySub(Arg1 := 3, Arg1 := 5) 

  ```


     Remove one of the duplicate specifications.
    
- You specified the same [argument](../../Glossary/vbe-glossary.md#argument) both by position and with a named argument, for example:
    
  ```vb
  Call MySub(1, Arg1 := 3) 

  ```


    Remove one of the duplicate specifications.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]