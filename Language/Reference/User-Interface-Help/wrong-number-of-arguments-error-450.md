---
title: Wrong number of arguments (Error 450)
keywords: vblr6.chm1000450
f1_keywords:
- vblr6.chm1000450
ms.prod: office
ms.assetid: 7a1af0b6-59f3-79c6-3167-3d94405ba23d
ms.date: 06/08/2019
localization_priority: Normal
---


# Wrong number of arguments (Error 450)

The number of [arguments](../../Glossary/vbe-glossary.md#argument) to a procedure must match the number of[parameters](../../Glossary/vbe-glossary.md#parameter) in the [procedure's](../../Glossary/vbe-glossary.md#procedure) definition. This error has the following causes and solutions:



- The number of arguments in the call to the procedure wasn't the same as the number of required arguments expected by the procedure. Check the argument list in the call against the procedure declaration or definition.
    
- You specified an index for a control that isn't part of a [control array](../../Glossary/vbe-glossary.md#control-array).
    
    The index specification is interpreted as an argument but neither an index nor an argument is expected, so the error occurs. Remove the index specification, or follow the procedure for creating a control array. Set the  **Index** property to a nonzero value in the control's property sheet or property window at [design time](../../Glossary/vbe-glossary.md#design-time).
    
- You tried to assign a value to a read-only [property](../../Glossary/vbe-glossary.md#property), or you tried to assign a value to a property for which no  **Property Let** procedure exists.
    
    Assigning a value to a property is the same as passing the value as an argument to the object's  **Property Let** procedure. Properly define the **Property Let** procedure; it must have one more argument than the corresponding **Property Get** procedure. If the property is meant to be read-only, you can't assign a value to it.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
