---
title: Comment object (Word)
keywords: vbawd10.chm2365
f1_keywords:
- vbawd10.chm2365
ms.prod: word
api_name:
- Word.Comment
ms.assetid: 0a2841f3-ca3c-8186-afab-f634ebd97d4c
ms.date: 06/08/2017
localization_priority: Normal
---


# Comment object (Word)

Represents a single comment. The  **Comment** object is a member of the **[Comments](Word.comments.md)** collection. The **Comments** collection includes comments in a selection, range or document.


## Remarks

Use  **Comments** (Index), where Index is the index number, to return a single **Comment** object. The index number represents the position of the comment in the specified selection, range, or document. The following example displays the author of the first comment in the active document.


```vb
MsgBox ActiveDocument.Comments(1).Author
```

Use the  **[Add](Word.Comments.Add.md)** method to add a comment at the specified range. The following example adds a comment immediately after the selection.




```vb
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Comments.Add Range:=Selection.Range, _ 
 Text:="review this"
```

Use the  **[Reference](Word.Comment.Reference.md)** property to return the reference mark associated with the specified comment. Use the **[Range](Word.Comment.Range.md)** property to return the text associated with the specified comment. The following example displays the text associated with the first comment in the active document.




```vb
MsgBox ActiveDocument.Comments(1).Range.Text
```


## Methods



|Name|
|:-----|
|[DeleteRecursively](Word.comment.deleterecursively.md)|
|[Edit](Word.Comment.Edit.md)|

## Properties



|Name|
|:-----|
|[Ancestor](Word.comment.ancestor.md)|
|[Application](Word.Comment.Application.md)|
|[Contact](Word.comment.contact.md)|
|[Creator](Word.Comment.Creator.md)|
|[Date](Word.Comment.Date.md)|
|[Done](Word.comment.done.md)|
|[Index](Word.Comment.Index.md)|
|[IsInk](Word.Comment.IsInk.md)|
|[Parent](Word.Comment.Parent.md)|
|[Range](Word.Comment.Range.md)|
|[Reference](Word.Comment.Reference.md)|
|[Replies](Word.comment.replies.md)|
|[Scope](Word.Comment.Scope.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
