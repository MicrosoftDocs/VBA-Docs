---
title: KeyBindings.Context property (Word)
keywords: vbawd10.chm160825354
f1_keywords:
- vbawd10.chm160825354
ms.prod: word
api_name:
- Word.KeyBindings.Context
ms.assetid: 8cdba82a-a4cc-f549-a3c5-4bfbb70578b6
ms.date: 06/08/2017
localization_priority: Normal
---


# KeyBindings.Context property (Word)

Returns an **Object** that represents the storage location of the specified key binding. Read-only.


## Syntax

_expression_. `Context`

_expression_ A variable that represents a '[KeyBindings](Word.keybindings.md)' collection.


## Remarks

This property can return a  **Document**, **Template**, or **Application** object. Built-in key assignments (for example, CTRL+I for **Italic**) return the **Application** object as the context. Any key bindings you add will return a **Document** or **Template** object, depending on the customization context in effect when the **KeyBinding** object was added.


## See also


[KeyBindings Collection Object](Word.keybindings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]