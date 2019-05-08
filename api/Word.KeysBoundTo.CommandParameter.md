---
title: KeysBoundTo.CommandParameter property (Word)
keywords: vbawd10.chm160890885
f1_keywords:
- vbawd10.chm160890885
ms.prod: word
api_name:
- Word.KeysBoundTo.CommandParameter
ms.assetid: de72887d-0970-05e5-84e2-4ba4c5c6ae45
ms.date: 06/08/2017
localization_priority: Normal
---


# KeysBoundTo.CommandParameter property (Word)

Returns the command parameter assigned to the specified shortcut key. Read-only  **String**.


## Syntax

_expression_. `CommandParameter`

_expression_ A variable that represents a '[KeysBoundTo](Word.keysboundto.md)' object.


## Remarks

For information about commands that take parameters, see the  **[Add](Word.KeyBindings.Add.md)** method. Use the **Command** property to return the command name assigned to the specified shortcut key.


## Example

This example assigns a shortcut key to the FontSize command, with a command parameter of 8. Use the CommandParameter property to display the command parameter along with the command name and key string.


```vb
Dim kbNew As KeyBinding 
 
Set kbNew = KeyBindings.Add(KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="FontSize", _ 
 KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyS), _ 
 CommandParameter:="8") 
MsgBox kbNew.Command & Chr$(32) & kbNew.CommandParameter _ 
 & vbCr & kbNew.KeyString
```


## See also


[KeysBoundTo Collection Object](Word.keysboundto.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]