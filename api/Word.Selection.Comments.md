---
title: Selection.Comments property (Word)
keywords: vbawd10.chm158662712
f1_keywords:
- vbawd10.chm158662712
ms.prod: word
api_name:
- Word.Selection.Comments
ms.assetid: 8f6fda0e-7070-eb42-3e1b-3a2a0654b330
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Comments property (Word)

Returns a **[Comments](Word.comments.md)** collection that represents all the comments in the specified. Read-only.


## Syntax

_expression_.**Comments**

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example adds a comment to the selected text.

```vb
ActiveDocument.ActiveWindow.View.ShowHiddenText = True 
Selection.Comments.Add Range:=Selection.Range, Text:="Approved"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]