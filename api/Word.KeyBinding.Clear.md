---
title: KeyBinding.Clear method (Word)
keywords: vbawd10.chm160956517
f1_keywords:
- vbawd10.chm160956517
ms.prod: word
api_name:
- Word.KeyBinding.Clear
ms.assetid: 7f53f149-71e9-e2ff-c261-31cd1f0668de
ms.date: 06/08/2017
localization_priority: Normal
---


# KeyBinding.Clear method (Word)

Removes the specified key binding from the  **KeyBindings** collection and resets a built-in command to its default key assignment.


## Syntax

_expression_.**Clear**

_expression_ A variable that represents a '[KeyBinding](Word.KeyBinding.md)' object.


## Example

This example removes the ALT+F1 key assignment from the Normal template.


```vb
CustomizationContext = NormalTemplateFindKey(BuildKeyCode(Arg1:=wdKeyAlt, Arg2:=wdKeyF1)).Clear
```


## See also


[KeyBinding Object](Word.KeyBinding.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]