---
title: Comment.IsInk property (Word)
keywords: vbawd10.chm154993652
f1_keywords:
- vbawd10.chm154993652
ms.prod: word
api_name:
- Word.Comment.IsInk
ms.assetid: 57204e17-cf5a-d006-0738-b1f1ef62632f
ms.date: 06/08/2017
localization_priority: Normal
---


# Comment.IsInk property (Word)

Returns a  **Boolean** that represents whether a comment is a handwritten comment.


## Syntax

_expression_. `IsInk`

 _expression_ An expression that returns a '[Comment](Word.Comment.md)' object.


## Example

The following example removes all handwritten comments from the active document.


```vb
Dim objComment As Comment 
 
For Each objComment In ActiveDocument.Comments 
 If objComment.IsInk = True Then 
 objComment.Delete 
 End If 
Next
```


## See also


[Comment Object](Word.Comment.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]