---
title: CommentThreaded.AddReply method (Excel)
keywords: 
f1_keywords:
- 
ms.prod: excel
api_name:
- Excel.CommentThreaded.AddReply
ms.assetid: 
ms.date: 05/08/2019
localization_priority: Normal
---


# CommentThreaded.AddReply method (Excel)

If this comment is a thread/parent, add a reply to Replies.  
If this comment is a child/reply, then add the reply to the parent’s Replies.  


## Syntax

_expression_. `AddReply`( `_Text_` )

_expression_ A variable that represents a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Parameters


|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Optional| **Variant**|The reply's text.|

## Return value

CommentThreaded


## Example

This example adds a reply to the CommentThreaded on cell E5 on worksheet one.


```vb
Worksheets(1).Range("E5").CommentThreaded.AddReply "Current Sales"
```


## See also

**[CommentThreaded](Excel.CommentThreaded.md)**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
