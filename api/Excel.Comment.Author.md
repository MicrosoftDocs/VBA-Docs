---
title: Comment.Author property (Excel)
keywords: vbaxl10.chm516073
f1_keywords:
- vbaxl10.chm516073
ms.prod: excel
api_name:
- Excel.Comment.Author
ms.assetid: ac964a80-1646-41a0-8b3a-941c800395e7
ms.date: 04/23/2019
localization_priority: Normal
---


# Comment.Author property (Excel)

Returns the author of the comment. Read-only **String**.

## Syntax

_expression_.**Author**

_expression_ A variable that represents a **[Comment](Excel.Comment.md)** object.


## Example

This example deletes all comments added by author Jean Selva on the active sheet.

```vb
For Each c in ActiveSheet.Comments 
 If c.Author = "Jean Selva" Then c.Delete 
Next
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]