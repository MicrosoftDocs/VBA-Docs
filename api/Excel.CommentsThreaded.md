---
title: CommentsThreaded object (Excel)
keywords: vbaxl10.chm1007072
f1_keywords:
- vbaxl10.chm1007072
ms.prod: excel
api_name:
- Excel.CommentsThreaded
ms.date: 06/21/2019
localization_priority: Normal
---


# CommentsThreaded object (Excel)

A collection of top-level **[CommentThreaded](Excel.CommentThreaded.md)** objects in a **Worksheet**, or a collection of replies in a single threaded comment.


## Remarks

Each threaded comment is represented by a **CommentThreaded** object.


## Example

Use the **[CommentsThreaded](excel.worksheet.commentsThreaded.md)** property of the **Worksheet** object to return the **CommentsThreaded** collection. The following example updates the text of all the threaded comments on worksheet one.

```vb
Set cmt = Worksheets(1).CommentsThreaded 
For Each c In cmt 
 c.Text "Updated Comment"
Next
```

<br/>

Use the **[AddCommentThreaded](Excel.Range.AddCommentThreaded.md)** method of the **Range** object to add a threaded comment to a range. The following example adds a threaded comment to cell E5 on worksheet one.

```vb
Worksheets(1).Range("e5").AddCommentThreaded("This is a Threaded Comment")
```

<br/>

Use **CommentsThreaded** (_index_), where _index_ is the threaded comment number, to return a single threaded comment from the **CommentsThreaded** collection. The following example updates the text of threaded comment two on worksheet one.

```vb
Worksheets(1).CommentsThreaded(2).Text "Updated Text"
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