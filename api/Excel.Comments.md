---
title: Comments object (Excel)
keywords: vbaxl10.chm513072
f1_keywords:
- vbaxl10.chm513072
ms.prod: excel
api_name:
- Excel.Comments
ms.assetid: f43bf021-1e46-10cf-09bf-070fc6a2c81a
ms.date: 03/29/2019
localization_priority: Normal
---


# Comments object (Excel)

A collection of cell comments.


## Remarks

Each comment is represented by a **[Comment](Excel.Comment.md)** object.


## Example

Use the **[Comments](excel.worksheet.comments.md)** property of the **Worksheet** object to return the **Comments** collection. The following example hides all the comments on worksheet one.

```vb
Set cmt = Worksheets(1).Comments 
For Each c In cmt 
 c.Visible = False 
Next
```

<br/>

Use the **[AddComment](Excel.Range.AddComment.md)** method of the **Range** object to add a comment to a range. The following example adds a comment to cell E5 on worksheet one.

```vb
With Worksheets(1).Range("e5").AddComment 
 .Visible = False 
 .Text "reviewed on " & Date 
End With
```

<br/>

Use **Comments** (_index_), where _index_ is the comment number, to return a single comment from the **Comments** collection. The following example hides comment two on worksheet one.

```vb
Worksheets(1).Comments(2).Visible = False
```



## Methods

- [Item](Excel.Comments.Item.md)

## Properties

- [Application](Excel.Comments.Application.md)
- [Count](Excel.Comments.Count.md)
- [Creator](Excel.Comments.Creator.md)
- [Parent](Excel.Comments.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]