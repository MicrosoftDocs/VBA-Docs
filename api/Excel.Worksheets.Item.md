---
title: Worksheets.Item property (Excel)
keywords: vbaxl10.chm470078
f1_keywords:
- vbaxl10.chm470078
ms.prod: excel
api_name:
- Excel.Worksheets.Item
ms.assetid: 66099ca2-54a0-f8ae-58ab-07791bbf5e7c
ms.date: 05/18/2019
localization_priority: Normal
---


# Worksheets.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Worksheets](Excel.Worksheets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Remarks

For more information about returning a single member of a collection, see [Returning an object from a collection](../excel/Concepts/Workbooks-and-Worksheets/returning-an-object-from-a-collection-excel.md).


## Example

**Item** is the default member for a collection. For example, the following two lines of code are equivalent.

```vb
ActiveWorkbook.Worksheets.Item(1) 
ActiveWorkbook.Worksheets(1)
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
