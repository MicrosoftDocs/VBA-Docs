---
title: Range.AddCommentThreaded method (Excel)
keywords: vbaxl10.chm144259
f1_keywords:
- vbaxl10.chm144259
ms.prod: excel
api_name:
- Excel.Range.AddCommentThreaded
ms.date: 05/15/2019
localization_priority: Normal
---


# Range.AddCommentThreaded method (Excel)

Adds a new modern threaded comment to the range if no comment already exists. 


## Syntax

_expression_.**AddCommentThreaded** (_Text_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The comment text.|

## Return value

**CommentThreaded**


## Example

This example adds a threaded comment to cell E5 on worksheet one.

```vb
Worksheets(1).Range("E5").AddCommentThreaded "Current Sales"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
