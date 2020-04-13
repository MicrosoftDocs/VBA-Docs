---
title: Application.KeyBindings property (Word)
keywords: vbawd10.chm158335045
f1_keywords:
- vbawd10.chm158335045
ms.prod: word
api_name:
- Word.Application.KeyBindings
ms.assetid: 68e08a9a-6547-f722-078e-b603b9f3e9cb
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.KeyBindings property (Word)

Returns a  **[KeyBindings](Word.keybindings.md)** collection that represents customized key assignments, which include a key code, a key category, and a command.


## Syntax

_expression_. `KeyBindings`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example assigns the CTRL+ALT+W key combination to the **FileClose** command. This keyboard customization is saved in the Normal template.


```vb
CustomizationContext = NormalTemplate 
KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, _ 
 wdKeyW), KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="FileClose"
```

This example inserts the command name and key combination string for each item in the KeyBindings collection.




```vb
Dim kbLoop As KeyBinding 
 
CustomizationContext = NormalTemplate 
For Each kbLoop In KeyBindings 
 Selection.InsertAfter kbLoop.Command & vbTab _ 
 & kbLoop.KeyString & vbCr 
 Selection.Collapse Direction:=wdCollapseEnd 
Next kbLoop
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]