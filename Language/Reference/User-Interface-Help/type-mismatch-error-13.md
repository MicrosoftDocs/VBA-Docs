---
title: Type mismatch (Error 13)
keywords: vblr6.chm1011290
f1_keywords:
- vblr6.chm1011290
ms.prod: office
ms.assetid: cbc7e902-b468-c335-5620-1ff9a2026b9b
ms.date: 08/14/2019
localization_priority: Normal
---


# Type mismatch (Error 13)

Visual Basic is able to convert and coerce many values to accomplish [data type](../../Glossary/vbe-glossary.md#data-type) assignments that weren't possible in earlier versions.

However, this error can still occur and has the following causes and solutions:

- **Cause:** The [variable](../../Glossary/vbe-glossary.md#variable) or [property](../../Glossary/vbe-glossary.md#property) isn't of the correct type. For example, a variable that requires an integer value can't accept a string value unless the whole string can be recognized as an integer.
    
 **Solution:** Try to make assignments only between compatible [data types](../../Glossary/vbe-glossary.md#data-type). For example, an **Integer** can always be assigned to a **Long**, a **Single** can always be assigned to a **Double**, and any type (except a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type)) can be assigned to a **Variant**.
    
- **Cause:** An object was passed to a [procedure](../../Glossary/vbe-glossary.md#procedure) that is expecting a single property or value.
    
 **Solution:** Pass the appropriate single property or call a [method](../../Glossary/vbe-glossary.md#method) appropriate to the object.
    
- **Cause:** A [module](../../Glossary/vbe-glossary.md#module) or [project](../../Glossary/vbe-glossary.md#project) name was used where an [expression](../../Glossary/vbe-glossary.md#expression) was expected, for example:
    
  ```vb
    Debug.Print MyModule 
  ```

 **Solution:** Specify an expression that can be displayed.
    
- **Cause:** You attempted to mix traditional Basic error handling with **Variant** values having the **Error** subtype (10, **vbError**), for example:
    
  ```vb
    Error CVErr(n) 
  ```

 **Solution:** To regenerate an error, you must map it to an intrinsic Visual Basic or a user-defined error, and then generate that error.
    
- **Cause:** A **CVErr** value can't be converted to **Date**. For example:
    
  ```vb
    MyVar = CDate(CVErr(9)) 
  ```

 **Solution:** Use a **Select Case** statement or some similar construct to map the return of **CVErr** to such a value.
    
- **Cause:** At [run time](../../Glossary/vbe-glossary.md#run-time), this error typically indicates that a **Variant** used in an expression has an incorrect subtype, or a **Variant** containing an [array](../../Glossary/vbe-glossary.md#array) appears in a **Print #** statement.
    
 **Solution:** To print arrays, create a loop that displays each element individually.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Add-ins note](~/includes/addinsnote.md)]

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
