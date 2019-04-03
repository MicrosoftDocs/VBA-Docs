---
title: AutoCorrect.CorrectTableCells property (Word)
keywords: vbawd10.chm155779091
f1_keywords:
- vbawd10.chm155779091
ms.prod: word
api_name:
- Word.AutoCorrect.CorrectTableCells
ms.assetid: 8bb5dfdd-9c54-b49e-609f-18b4d8b556ee
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.CorrectTableCells property (Word)

 **True** to automatically capitalize the first letter of table cells. Read/write **Boolean**.


## Syntax

_expression_. `CorrectTableCells`

 _expression_ An expression that returns an '[AutoCorrect](Word.AutoCorrect.md)' object.


## Example

This example disables automatic capitalization of the first letter typed within table cells.


```vb
Sub AutoCorrectFirstLetterOfTableCells() 
 Application.AutoCorrect.CorrectTableCells = False 
End Sub
```


## See also


[AutoCorrect Object](Word.AutoCorrect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]