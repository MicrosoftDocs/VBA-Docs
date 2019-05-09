---
title: PivotTable.RowAxisLayout method (Excel)
keywords: vbaxl10.chm235166
f1_keywords:
- vbaxl10.chm235166
ms.prod: excel
api_name:
- Excel.PivotTable.RowAxisLayout
ms.assetid: 41a8a3bb-252a-7598-b559-d75dc1e10bc1
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.RowAxisLayout method (Excel)

This method is used for simultaneously setting layout options for all existing PivotFields.


## Syntax

_expression_.**RowAxisLayout** (_RowLayout_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RowLayout_|Required| **[XlLayoutRowType](excel.xllayoutrowtype.md)** |Specifies the type of layout row. Can be **xlCompactRow**, **xlTabularRow**, or **xlOutlineRow**.|

## Remarks

This method is atomic so it makes sure that if layout options cannot be set on any PivotField, the layout options of none of the fields will change, and no change is made to the PivotTable.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]