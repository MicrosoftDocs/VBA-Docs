---
title: Charts.Add2 method (Excel)
keywords: vbaxl10.chm218076
f1_keywords:
- vbaxl10.chm218076
ms.prod: excel
ms.assetid: bfd7d614-a640-dfdc-ebc5-3d0682f2c839
ms.date: 04/20/2019
localization_priority: Normal
---


# Charts.Add2 method (Excel)

Inserts a chart directly onto the grid.


## Syntax

_expression_.**Add2** (_Before_, _After_, _Count_, _NewLayout_)

_expression_ A variable that represents a **[Charts](Excel.Charts.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Before_|Optional|**Variant**|An object that specifies the sheet before which the new sheet is added.|
| _After_|Optional|**Variant**|An object that specifies the sheet after which the new sheet is added.|
| _Count_|Optional|**Variant**|The number of sheets to be added. The default value is one.|
| _NewLayout_|Optional|**Variant**|If **NewLayout** is **True**, the chart is inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).|

## Return value

**CHART**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]