---
title: CommentThreaded object (Excel)
keywords: vbaxl10.chm1009072
f1_keywords:
- vbaxl10.chm1009072
ms.prod: excel
api_name:
- Excel.CommentThreaded
ms.date: 05/15/2019
localization_priority: Normal
---


# CommentThreaded object (Excel)

Represents a cell's threaded comment. This object can represent both a top-level comment or its replies.


## Remarks

The **CommentThreaded** object is a member of the **[CommentsThreaded](Excel.CommentsThreaded.md)** collection.


## Example

Use the **[CommentThreaded](Excel.Range.CommentThreaded.md)** property of the **Range** object to return a **CommentThreaded** object. The following example changes the text in the threaded comment in cell E5.

```vb
Worksheets(1).Range("E5").CommentThreaded.Text "reviewed on " & Date
```

<br/>

Use **CommentsThreaded** (_index_), where _index_ is the threaded comment number, to return a single threaded comment from the **CommentsThreaded** collection. The following example updates the text of threaded comment two on worksheet one.

```vb
Worksheets(1).CommentsThreaded(2).Text "reviewed on " & Date
```

<br/>

Use the **[AddCommentThreaded](Excel.Range.AddCommentThreaded.md)** method of the **Range** object to add a comment to a range. The following example adds a threaded comment to cell E5 on worksheet one.

```vb
Worksheets(1).Range("E5").AddCommentThreaded "Current Sales"
```


## Methods

- [AddReply](Excel.CommentThreaded.AddReply.md)
- [Delete](Excel.CommentThreaded.Delete.md)
- [Next](Excel.CommentThreaded.Next.md)
- [Previous](Excel.CommentThreaded.Previous.md)
- [Text](Excel.CommentThreaded.Text.md)

## Properties

- [Application](Excel.CommentThreaded.Application.md)
- [Author](Excel.CommentThreaded.Author.md)
- [Creator](Excel.CommentThreaded.Creator.md)
- [Date](Excel.CommentThreaded.Date.md)
- [Parent](Excel.CommentThreaded.Parent.md)
- [Replies](Excel.CommentThreaded.Replies.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
