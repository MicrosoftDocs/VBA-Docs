---
title: User-defined type not defined (VBA)
keywords: vblr6.chm1011292
f1_keywords:
- vblr6.chm1011292
ms.prod: office
ms.assetid: 60e0da5e-c498-7a2f-46c6-c09d59fc607a
ms.date: 12/27/2018
localization_priority: Priority
---


# User-defined type not defined

You can create your own [data types](../Glossary/vbe-glossary.md#data-type) in Visual Basic, but they must be defined first in a **Type...End Type** statement or in a properly registered [object library](../Glossary/vbe-glossary.md#object-library) or [type library](../Glossary/vbe-glossary.md#type-library). This error has the following causes and solutions:

- You tried to declare a [variable](../Glossary/vbe-glossary.md#variable) or [argument](../Glossary/vbe-glossary.md#argument) with an undefined data type or you specified an unknown [class](../Glossary/vbe-glossary.md#class) or object.
    
  Use the **Type** statement in a [module](../Glossary/vbe-glossary.md#module) to define a new data type. If you are trying to create a reference to a class, the class must be visible to the [project](../Glossary/vbe-glossary.md#project). If you are referring to a class in your program, you must have a [class module](../Glossary/vbe-glossary.md#class-module) of the specified name in your project. Check the spelling of the type name or name of the object.
    
- The type you want to declare is in another module but has been declared **Private**. Move the definition of the type to a [standard module](../Glossary/vbe-glossary.md#standard-module) where it can be **Public**.
    
- The type is a valid type, but the object library or type library in which it is defined isn't registered in Visual Basic. Display the **References** dialog box, and then select the appropriate object library or type library. For example, if you don't check the **Data Access Object** in the **References** dialog box, types like Database, Recordset, and TableDef aren't recognized and references to them in code cause this error.
    
For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

## See also

- [Visual Basic how-to topics](../reference/user-interface-help/visual-basic-how-to-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
