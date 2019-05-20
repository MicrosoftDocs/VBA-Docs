---
title: Global.BuildKeyCode method (Word)
keywords: vbawd10.chm163119420
f1_keywords:
- vbawd10.chm163119420
ms.prod: word
api_name:
- Word.Global.BuildKeyCode
ms.assetid: dc9870a9-0c0d-5985-e3fc-79c5a1b467c6
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.BuildKeyCode method (Word)

Returns a unique number for the specified key combination.


## Syntax

_expression_. `BuildKeyCode` (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a '[Global](Word.Global.md)' object. Optional.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **[WdKey](Word.WdKey.md)**|A key you specify by using one of the  **WdKey** constants.|
| _Arg2_|Optional| **[WdKey](Word.WdKey.md)**|A key you specify by using one of the  **WdKey** constants.|
| _Arg3_|Optional| **[WdKey](Word.WdKey.md)**|A key you specify by using one of the  **WdKey** constants.|
| _Arg4_|Optional| **[WdKey](Word.WdKey.md)**|A key you specify by using one of the  **WdKey** constants.|

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


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]