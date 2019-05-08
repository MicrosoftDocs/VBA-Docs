---
title: Range.AddCommentThreaded method (Excel)
keywords: 
f1_keywords:
- 
ms.prod: excel
api_name:
- Excel.Range.AddCommentThreaded
ms.assetid: 
ms.date: 05/08/2019
localization_priority: Normal
---


# Range.AddCommentThreaded method (Excel)

Adds a new modern CommentThreaded to the range if no comment already exists. 


## Syntax

_expression_. `AddCommentThreaded`( `_Text_` )

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters


|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Optional| **Variant**|The comment text.|

## Return value

CommentThreaded


## Example

This example adds a CommentThreaded to cell E5 on worksheet one.


```vb
Worksheets(1).Range("E5").AddCommentThreaded "Current Sales"
```


## See also


[Range Object](Excel.Range(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
