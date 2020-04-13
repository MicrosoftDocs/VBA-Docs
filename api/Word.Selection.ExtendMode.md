---
title: Selection.ExtendMode property (Word)
keywords: vbawd10.chm158663062
f1_keywords:
- vbawd10.chm158663062
ms.prod: word
api_name:
- Word.Selection.ExtendMode
ms.assetid: 7b12cf8b-9be1-6ebc-de96-e7734eaad3b6
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ExtendMode property (Word)

 **True** if Extend mode is active. Read/write **Boolean**.


## Syntax

_expression_. `ExtendMode`

 _expression_ An expression that returns a **[Selection](Word.Selection.md)** object.


## Remarks

When Extend mode is active, the Extend argument of the following methods is **True** by default: **[EndKey](Word.Selection.EndKey.md)**, **[HomeKey](Word.Selection.HomeKey.md)**, **[MoveDown](Word.Selection.MoveDown.md)**, **[MoveLeft](Word.Selection.MoveLeft.md)**, **[MoveRight](Word.Selection.MoveRight.md)**, and **[MoveUp](Word.Selection.MoveUp.md)**. Also, the letters "EXT" appear on the status bar.

This property can only be set during run time; attempts to set it in Immediate mode are ignored. The Extend arguments of the **[EndOf](Word.Selection.EndOf.md)** and **[StartOf](Word.Selection.StartOf.md)** methods are not affected by this property.


## Example

This example moves to the beginning of the paragraph and selects the paragraph plus the next two sentences.


```vb
With Selection 
 .MoveUp Unit:=wdParagraph 
 .ExtendMode = True 
 .MoveDown Unit:=wdParagraph 
 .MoveRight Unit:=wdSentence, Count:=2 
End With
```

This example collapses the current selection, turns on Extend mode, and selects the current sentence.




```vb
With Selection 
 .Collapse 
 .ExtendMode = True 
 ' Select current word. 
 .Extend 
 ' Select current sentence. 
 .Extend 
End With
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]