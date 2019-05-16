---
title: CommentThreaded.Delete method (Excel)
keywords: vbaxl10.chm1010074
f1_keywords:
- vbaxl10.chm1010074
ms.prod: excel
api_name:
- Excel.CommentThreaded.Delete
ms.date: 05/15/2019
localization_priority: Normal
---


# CommentThreaded.Delete method (Excel)

Deletes the specified threaded comment and all replies associated with that comment (if any exist). 


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Example

This example deletes the threaded comment on cell E5 on worksheet one.

```vb
Worksheets(1).Range("E5").CommentThreaded.Delete
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]