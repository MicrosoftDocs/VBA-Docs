---
title: Chart.Move method (Excel)
keywords: vbaxl10.chm148079
f1_keywords:
- vbaxl10.chm148079
ms.prod: excel
api_name:
- Excel.Chart.Move
ms.assetid: ec8c8eae-17a8-20a0-a87c-81f31b21d735
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Move method (Excel)

Moves the chart to another location in the workbook.


## Syntax

_expression_.**Move** (_Before_, _After_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|The sheet before which the moved chart will be placed. You cannot specify _Before_ if you specify _After_.|
| _After_|Optional| **Variant**| The sheet after which the moved chart will be placed. You cannot specify _After_ if you specify _Before_.|

## Remarks

If you don't specify either _Before_ or _After_, Microsoft Excel creates a new workbook that contains the moved chart.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]