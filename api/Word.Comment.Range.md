---
title: Comment.Range property (Word)
keywords: vbawd10.chm154993643
f1_keywords:
- vbawd10.chm154993643
ms.prod: word
api_name:
- Word.Comment.Range
ms.assetid: 1a67e361-67ee-0fb1-ffe4-9e15aa73e2a2
ms.date: 06/08/2017
localization_priority: Normal
---


# Comment.Range property (Word)

Returns a  **Range** object that represents the contents of a comment.


## Syntax

 _expression_. `Range`

 _expression_ Required. A variable that represents a '[Comment](Word.Comment.md)' object.


## Example

This example changes the text of the first comment in the document.


```vb
With ActiveDocument.Comments(1).Range 
 .Delete 
 .InsertBefore "new comment text" 
End With
```


## See also


[Comment Object](Word.Comment.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]