---
title: KeyBinding.KeyCode property (Word)
keywords: vbawd10.chm160956422
f1_keywords:
- vbawd10.chm160956422
ms.prod: word
api_name:
- Word.KeyBinding.KeyCode
ms.assetid: 8ca07f1e-b60b-bc10-b1fe-cb0d7b890d33
ms.date: 06/08/2017
localization_priority: Normal
---


# KeyBinding.KeyCode property (Word)

Returns a unique number for the first key in the specified key binding. Read-only  **Long**.


## Syntax

_expression_. `KeyCode`

 _expression_ An expression that returns a '[KeyBinding](Word.KeyBinding.md)' object.


## Remarks

You create this number by using the **[BuildKeyCode](Word.Application.BuildKeyCode.md)** method when you are adding key bindings by using the **[Add](Word.KeyBindings.Add.md)** method of the **[KeyBindings](Word.keybindings.md)** object.


## Example

This example displays a message if the **KeyBindings** collection includes the ALT+CTRL+W key combination.


```vb
Dim lngCode As Long 
Dim kbLoop As KeyBinding 
 
CustomizationContext = NormalTemplate 
lngCode = BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyW) 
For Each kbLoop In KeyBindings 
 If lngCode = kbLoop.KeyCode Then 
 MsgBox kbLoop.KeyString & " is already in use" 
 End If 
Next kbLoop
```


## See also


[KeyBinding Object](Word.KeyBinding.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]