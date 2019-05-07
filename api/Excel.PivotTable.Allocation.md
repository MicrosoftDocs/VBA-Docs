---
title: PivotTable.Allocation property (Excel)
keywords: vbaxl10.chm235187
f1_keywords:
- vbaxl10.chm235187
ms.prod: excel
api_name:
- Excel.PivotTable.Allocation
ms.assetid: ac7bd537-97f0-f643-3e34-dd13e49ac149
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.Allocation property (Excel)

Returns or sets whether to run an **UPDATE CUBE** statement for each cell that is edited, or only when the user chooses to calculate changes when performing what-if analysis on a PivotTable based on an OLAP data source. Read/write.


## Syntax

_expression_.**Allocation**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

**[XlAllocation](Excel.XlAllocation.md)**


## Remarks

The **Allocation** property corresponds to the **Calculate with changes** setting in the **What-If Analysis Settings** dialog box. The default setting is **xlManualAllocation**, which corresponds to the **Manually (when selecting calculate PivotTable with changes)** setting.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]