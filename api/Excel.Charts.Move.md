---
title: Charts.Move method (Excel)
keywords: vbaxl10.chm217077
f1_keywords:
- vbaxl10.chm217077
api_name:
- Excel.Charts.Move
ms.assetid: 2f056384-6da5-4431-0458-a583e7f975d7
ms.date: 04/20/2019
ms.localizationpriority: medium
---


# Charts.Move method (Excel)

Moves the chart to another location in the workbook.


## Syntax

_expression_.**Move** (_Before_, _After_)

_expression_ A variable that represents a **[Charts](Excel.Charts.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Optional| **Variant**|The sheet before which the moved chart will be placed. You cannot specify _Before_ if you specify _After_.|
| _After_|Optional| **Variant**| The sheet after which the moved chart will be placed. You cannot specify _After_ if you specify _Before_.|

## Remarks

If you don't specify either _Before_ or _After_, Microsoft Excel creates a new workbook that contains the moved chart.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]