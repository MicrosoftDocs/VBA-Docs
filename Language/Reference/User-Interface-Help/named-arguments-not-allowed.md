---
title: Named arguments not allowed
keywords: vblr6.chm1040129
f1_keywords:
- vblr6.chm1040129
ms.prod: office
ms.assetid: 886826a2-6d43-ec66-da42-7528a8470f9f
ms.date: 06/08/2017
localization_priority: Normal
---


# Named arguments not allowed

[Named arguments](../../Glossary/vbe-glossary.md#named-argument) aren't permitted in all situations. This error has the following causes and solutions:



- You tried to specify a named argument as an [array](../../Glossary/vbe-glossary.md#array) index, for example:
    
  ```vb
  MyVar = MyArray(MyNamedArg := 1) 

  ```


    Use an ordinary [variable](../../Glossary/vbe-glossary.md#variable) or constant[expression](../../Glossary/vbe-glossary.md#expression) as an array index.
    
- You tried to specify a named argument with an object, for example:
    
  ```vb
  MyVar = MyObject(MyNamedArg := 1) 

  ```


     Use a variable or constant expression if the object requires an [argument](../../Glossary/vbe-glossary.md#argument). For example, if the default for an object is a [method](../../Glossary/vbe-glossary.md#method), the object's name represents the default method. If it needs arguments, specify them positionally.
    
- You tried to specify a named argument with an external name:
    
  ```vb
  MyVar = [MyName](MyNamedArg := 1) 

  ```


     Use an ordinary variable or constant expression if the external name needs an argument.
    
- You tried to specify a named argument with a data member of an object, for example:
    
  ```vb
  MyVar = [MyObject].MyProperty(MyNamedArg := 1) 

  ```


     Use an ordinary variable or constant expression if the data member needs an argument.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]