---
title: Comment object (Excel)
keywords: vbaxl10.chm515072
f1_keywords:
- vbaxl10.chm515072
ms.prod: excel
api_name:
- Excel.Comment
ms.assetid: 3627e9be-2a28-9dc5-c822-ad42857134e3
ms.date: 03/29/2019
localization_priority: Normal
---


# Comment object (Excel)

Represents a cell comment.


## Remarks

The **Comment** object is a member of the **[Comments](Excel.Comments.md)** collection.


## Example

Use the **[Comment](Excel.Range.Comment.md)** property of the **Range** object to return a **Comment** object. The following example changes the text in the comment in cell E5.

```vb
Worksheets(1).Range("E5").Comment.Text "reviewed on " & Date
```

<br/>

Use **Comments** (_index_), where _index_ is the comment number, to return a single comment from the **Comments** collection. The following example hides comment two on worksheet one.

```vb
Worksheets(1).Comments(2).Visible = False
```

<br/>

Use the **[AddComment](Excel.Range.AddComment.md)** method of the **Range** object to add a comment to a range. The following example adds a comment to cell E5 on worksheet one.

```vb
With Worksheets(1).Range("e5").AddComment 
 .Visible = False 
 .Text "reviewed on " & Date 
End With
```


## Methods

- [Delete](Excel.Comment.Delete.md)
- [Next](Excel.Comment.Next.md)
- [Previous](Excel.Comment.Previous.md)
- [Text](Excel.Comment.Text.md)

## Properties

- [Application](Excel.Comment.Application.md)
- [Author](Excel.Comment.Author.md)
- [Creator](Excel.Comment.Creator.md)
- [Parent](Excel.Comment.Parent.md)
- [Shape](Excel.Comment.Shape.md)
- [Visible](Excel.Comment.Visible.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
