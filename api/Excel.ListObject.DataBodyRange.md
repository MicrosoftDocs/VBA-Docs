---
title: ListObject.DataBodyRange property (Excel)
keywords: vbaxl10.chm734082
f1_keywords:
- vbaxl10.chm734082
ms.prod: excel
api_name:
- Excel.ListObject.DataBodyRange
ms.assetid: fe906555-d006-8220-d9f8-59636cca68d5
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.DataBodyRange property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the range of values, excluding the header row, in a table. Read-only.


## Syntax

_expression_.**DataBodyRange**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Example

This example selects the active data range in the list.

```vb
Worksheets("Sheet1").Activate 
ActiveSheet.ListObjects.Item(1).DataBodyRange.Select
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
