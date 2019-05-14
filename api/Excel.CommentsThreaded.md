---
title: CommentsThreaded object (Excel)
keywords:
f1_keywords:
-
ms.prod: excel
api_name:
- Excel.CommentsThreaded
ms.assetid:
ms.date: 05/08/2019
localization_priority: Normal
---


# Comments object (Excel)

A collection of top-level **CommentThreaded** objects in a **Worksheet** or a collection of replies in a single **CommentThread**.


## Remarks

Each CommentThread is represented by a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Example

Use the **[CommentsThreaded](excel.worksheet.commentsThreaded.md)** property of the **Worksheet** object to return the **CommentsThreaded** collection. The following example updates the text all the CommentThreads on worksheet one.

```vb
Set cmt = Worksheets(1).CommentsThreaded 
For Each c In cmt 
 c.Text "Updated Comment"
Next
```

<br/>

Use the **[AddCommentThreaded](Excel.Range.AddCommentThreaded.md)** method of the **Range** object to add a CommentThreaded to a range. The following example adds a CommentThread to cell E5 on worksheet one.

```vb
With Worksheets(1).Range("e5").AddCommentThreaded
 .Text "reviewed on " & Date 
End With
```

<br/>

Use **CommentsThreaded** (_index_), where _index_ is the comment number, to return a single comment from the **CommentsThreaded** collection. The following example updates comment two's text on worksheet one.

```vb
Worksheets(1).Comments(2).Text "Updated Text"
```



## Methods

- [Item](Excel.CommentsThreaded.Item.md)

## Properties

- [Application](Excel.CommentsThreaded.Application.md)
- [Count](Excel.CommentsThreaded.Count.md)
- [Creator](Excel.CommentsThreaded.Creator.md)
- [Parent](Excel.CommentsThreaded.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]