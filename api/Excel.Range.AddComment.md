---
title: Range.AddComment method (Excel)
keywords: vbaxl10.chm144222
f1_keywords:
- vbaxl10.chm144222
ms.prod: excel
api_name:
- Excel.Range.AddComment
ms.assetid: 89bbacad-4655-bcc1-8010-2ab367cc7b31
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.AddComment method (Excel)

Adds a comment to the range.


## Syntax

_expression_. `AddComment`( `_Text_` )

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Optional| **Variant**|The comment text.|

## Return value

Comment


## Example

This example adds a comment to cell E5 on worksheet one.


```vb
Worksheets(1).Range("E5").AddComment "Current Sales"
```


## See also


[Range Object](Excel.Range(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]