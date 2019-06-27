---
title: CommentThreaded.AddReply method (Excel)
keywords: vbaxl10.chm1010073
f1_keywords:
- vbaxl10.chm1010073
ms.prod: excel
api_name:
- Excel.CommentThreaded.AddReply
ms.date: 06/27/2019
localization_priority: Normal
---


# CommentThreaded.AddReply method (Excel)

If the comment is a top-level comment, it will add a reply to its replies collection.

If this comment is a reply, it will add a reply to its Parent's replies collection.


## Syntax

_expression_.**AddReply** (_Text_)

_expression_ A variable that represents a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Optional| **Variant**|The reply's text.|

## Return value

**CommentThreaded**


## Example

This example adds a reply to the threaded comment on cell E5 on worksheet one.

```vb
Worksheets(1).Range("E5").CommentThreaded.AddReply "Current Sales"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
