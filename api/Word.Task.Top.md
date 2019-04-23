---
title: Task.Top property (Word)
keywords: vbawd10.chm159514626
f1_keywords:
- vbawd10.chm159514626
ms.prod: word
api_name:
- Word.Task.Top
ms.assetid: d6777e38-ce29-da8b-5bab-52cf3f022703
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.Top property (Word)

Returns or sets the vertical position, in [points](../language/glossary/vbe-glossary.md#point), of the specified window. Read/write  **Long**.


## Syntax

_expression_.**Top**

_expression_ Required. A variable that represents a '[Task](Word.Task.md)' object.


## Example

This example starts the Calculator and positions its window 100 points from the top of the screen.


```vb
Shell "Calc.exe" 
With Tasks("Calculator") 
 .WindowState = wdWindowStateNormal 
 .Top = 100 
End With
```


## See also


[Task Object](Word.Task.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]