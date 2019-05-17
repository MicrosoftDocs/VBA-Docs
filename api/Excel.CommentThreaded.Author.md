---
title: CommentThreaded.Author property (Excel)
keywords: vbaxl10.chm1010077
f1_keywords:
- vbaxl10.chm1010077
ms.prod: excel
api_name:
- Excel.CommentThread.Author
ms.date: 05/15/2019
localization_priority: Normal
---


# CommentThreaded.Author property (Excel)

Returns the **[Author](Excel.Author.md)** object that represents the author of the specified **CommentThreaded** object. Read-only.

## Syntax

_expression_.**Author**

_expression_ A variable that represents a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Example

This example deletes all threaded comments added by author Jean Selva on the active sheet.

```vb
For Each c in ActiveSheet.CommentsThreaded
 If c.Author.Name = "Jean Selva" Then c.Delete 
Next
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]