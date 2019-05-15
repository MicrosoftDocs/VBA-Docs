---
title: CommentThreaded.AddReply method (Excel)
ms.prod: excel
api_name:
- Excel.CommentThreaded.AddReply
ms.date: 05/15/2019
localization_priority: Normal
---


# CommentThreaded.AddReply method (Excel)

If this comment is a thread/parent, adds a reply to Replies. 

If this comment is a child/reply, adds the reply to the parent’s Replies.  


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
