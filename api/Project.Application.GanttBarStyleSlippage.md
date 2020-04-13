---
title: Application.GanttBarStyleSlippage method (Project)
keywords: vbapj.chm84
f1_keywords:
- vbapj.chm84
ms.prod: project-server
api_name:
- Project.Application.GanttBarStyleSlippage
ms.assetid: 2c5ec6cd-d588-a43a-7b06-8338ecd8ae6e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GanttBarStyleSlippage method (Project)

Shows or hides slippage for the specified baseline on Gantt bars of the active view.


## Syntax

_expression_. `GanttBarStyleSlippage`( `_Baseline_`, `_Show_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Baseline_|Required|**Integer**|Specifies the baseline number. Valid values are 0 through 10.|
| _Show_|Required|**Boolean**|If  **True**, show the baseline slippage. If **False**, hide the baseline slippage.|

## Return value

 **Boolean**


## Remarks

On the Ribbon, the **GanttBarStyleSlippage** method corresponds to the **Slippage** drop-down list in the **Bar Styles** group on the **Format** tab for **Gantt Chart Tools**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]