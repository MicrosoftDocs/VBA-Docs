---
title: PivotTable.SubtotalLocation method (Excel)
keywords: vbaxl10.chm235167
f1_keywords:
- vbaxl10.chm235167
ms.prod: excel
api_name:
- Excel.PivotTable.SubtotalLocation
ms.assetid: df2655d8-9e5f-e9d2-ba88-f92a1d843dfb
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotTable.SubtotalLocation method (Excel)

This method changes the subtotal location for all existing PivotFields. Changing the subtotal location has an immediate visual effect only for fields in outline form, but it will be set for fields in tabular form as well. 


## Syntax

_expression_. `SubtotalLocation`( `_Location_` )

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Location_|Required| **xlSubtototalLocationType**|xlSubtotalLocationType can be either  **xlAtTop** or **xlAtBottom**.|

## Remarks

The  **SubtotalLocation** method sets the **LayoutSubtotalLocation** property for all existing PivotFields automatically.


## See also


[PivotTable Object](Excel.PivotTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]