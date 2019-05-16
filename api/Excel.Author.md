---
title: Author object (Excel)
keywords: vbaxl10.chm1011072
f1_keywords:
- vbaxl10.chm1011072
ms.prod: excel
api_name:
- Excel.Author
ms.date: 05/15/2019
localization_priority: Normal
---


# Author object (Excel)

Represents the author of the **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Remarks

Use the **[Author](Excel.CommentThreaded.Author.md)** property of the **CommentThreaded** object to return the **Author** object. 

## Example

The following example shows how to get the author's name from the threaded comment on cell A1 on worksheet one.

```vb
Worksheets(1).Range("A1").CommentThreaded.Author.Name
```


## Properties

- [Application](Excel.Author.Application.md)
- [Creator](Excel.Author.Creator.md)
- [Name](Excel.Author.Name.md)
- [Parent](Excel.Author.Parent.md)
- [ProviderID](Excel.Author.ProviderID.md)
- [UserID](Excel.Author.UserID.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]