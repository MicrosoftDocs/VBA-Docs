---
title: Application.BuildKeyCode method (Word)
keywords: vbawd10.chm158335292
f1_keywords:
- vbawd10.chm158335292
ms.prod: word
api_name:
- Word.Application.BuildKeyCode
ms.assetid: 0bbc515d-5f39-fed8-2b86-99979af928a9
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BuildKeyCode method (Word)

Returns a unique number for the specified key combination.


## Syntax

_expression_. `BuildKeyCode` (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **[WdKey](Word.WdKey.md)**|A key you specify by using one of the **WdKey** constants.|
| _Arg2 - Arg4_|Optional| **[WdKey](Word.WdKey.md)**|A key you specify by using one of the **WdKey** constants.|

## Example

This example assigns the ALT + F1 key combination to the Organizer command.


```vb
CustomizationContext = NormalTemplate 
KeyBindings.Add KeyCode:=BuildKeyCode(Arg1:=wdKeyAlt, _ 
 Arg2:=wdKeyF1), KeyCategory:=wdKeyCategoryCommand, _ 
 Command:="Organizer"
```

This example removes the ALT+F1 key assignment from the Normal template.




```vb
CustomizationContext = NormalTemplate 
FindKey(BuildKeyCode(Arg1:=wdKeyAlt, Arg2:=wdKeyF1)).Clear
```

This example displays the command assigned to the F1 key.




```vb
CustomizationContext = NormalTemplate 
MsgBox FindKey(BuildKeyCode(Arg1:=wdKeyF1)).Command
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]