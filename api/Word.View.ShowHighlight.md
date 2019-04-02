---
title: View.ShowHighlight property (Word)
keywords: vbawd10.chm161808397
f1_keywords:
- vbawd10.chm161808397
ms.prod: word
api_name:
- Word.View.ShowHighlight
ms.assetid: ec0a5e47-f792-742b-654c-2aa137ab3ff1
ms.date: 06/08/2017
localization_priority: Normal
---


# View.ShowHighlight property (Word)

 **True** if highlight formatting is displayed and printed with a document. Read/write **Boolean**.


## Syntax

_expression_. `ShowHighlight`

 _expression_ An expression that returns a '[View](Word.View.md)' object.


## Example

This example toggles the display of highlighting in the active document.


```vb
ActiveDocument.ActiveWindow.View.ShowHighlight = _ 
 Not ActiveDocument.ActiveWindow.View.ShowHighlight
```

This example prints the active document without highlight formatting.




```vb
With ActiveDocument 
 .ActiveWindow.View.ShowHighlight = False 
 .PrintOut 
End With
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]