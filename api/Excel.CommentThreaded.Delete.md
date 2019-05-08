---
title: Comment.Delete method (Excel)
keywords:
f1_keywords:
-
ms.prod: excel
api_name:
- Excel.CommentThreaded.Delete
ms.assetid:
ms.date: 05/08/2019
localization_priority: Normal
---


# CommentThreaded.Delete method (Excel)

Deletes the specified comment and all replies associated with that comment (if any exist). 


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Example

This example deletes the CommentThreaded on cell E5 on worksheet one.


```vb
Worksheets(1).Range("E5").CommentThreaded.Delete
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]