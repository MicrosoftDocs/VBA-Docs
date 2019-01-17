---
title: View.ShowHyphens property (Word)
keywords: vbawd10.chm161808402
f1_keywords:
- vbawd10.chm161808402
ms.prod: word
api_name:
- Word.View.ShowHyphens
ms.assetid: 2294ea01-1ddc-0d29-4fa4-a5285d6d7cfb
ms.date: 06/08/2017
localization_priority: Normal
---


# View.ShowHyphens property (Word)

 **True** if optional hyphens are displayed. An optional hyphen indicates where to break a word when it falls at the end of a line. Read/write **Boolean**.


## Syntax

 _expression_. `ShowHyphens`

 _expression_ An expression that returns a '[View](Word.View.md)' object.


## Example

This example inserts an optional hyphen before the selection and then displays optional hyphens in the active window.


```vb
Selection.InsertBefore Chr(31) 
ActiveDocument.ActiveWindow.View.ShowHyphens = True
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]