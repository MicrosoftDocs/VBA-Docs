---
title: CommentThreaded.Author property (Excel)
keywords:
f1_keywords:
-
ms.prod: excel
api_name:
- Excel.CommentThread.Author
ms.assetid:
ms.date: 05/08/2019
localization_priority: Normal
---


# CommentThreaded.Author property (Excel)

Returns the **Author** object that represents the author of the specified CommentThreaded. Read-only.

## Syntax

_expression_.**Author**

_expression_ A variable that represents a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Example

This example deletes all CommentThreaded who's author's name is Jean Selva on the active sheet.

```vb
For Each c in ActiveSheet.CommentsThreaded
 If c.Author.Name = "Jean Selva" Then c.Delete 
Next
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]