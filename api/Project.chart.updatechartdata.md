---
title: Chart.UpdateChartData method (Project)
keywords: vbapj.chm131637
f1_keywords:
- vbapj.chm131637
ms.prod: project-server
ms.assetid: ecdef74d-480c-05a7-757c-a5c2e3e7359c
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.UpdateChartData method (Project)

Updates the specified Project data on a chart.

## Syntax

_expression_.**UpdateChartData** (_Task_, _Timephased_, _GroupName_, _FilterName_, _LabelField_, _OutlineLevel_, _SafeArrayOfPjField_, _SafeArrayOfPjTimescaledData_, _TimeScaleUnit_, _TimescaleUnitCount_, _StartDate_, _FinishDate_)

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Task_|Required|**Boolean**|**True** to update the task data; otherwise, **False**.|
| _Timephased_|Required|**Boolean**|**True** to update the timephased data; otherwise, **False**.|
| _GroupName_|Optional|**String**|The name of the **[Group2](Project.Group2.md)** object (a group of tasks or resources) for the update.|
| _FilterName_|Optional|**String**|The name of the **[Filter](Project.Filter.md)** object for the update.|
| _LabelField_|Optional|**PjField**|Specifies the field for the update. Can be one of the **[PjField](Project.PjField.md)** constants.|
| _OutlineLevel_|Optional|**Integer**|Specifies the task outline level for the update. The default value is -1, which is all outline levels.|
| _SafeArrayOfPjField_|Optional|**Variant**|Specifies an array of fields for the update, where each item in the array can be a **[PjField](Project.PjField.md)** constant.|
| _SafeArrayOfPjTimescaledData_|Optional|**Variant**|Specifies an array of timescaled data for the update, where each item in the array can be a **[PjTimescaledData](Project.PjTimescaledData.md)** constant.|
| _TimeScaleUnit_|Optional|**PjTimescaleUnit**|Specifies a timescale unit for the update. Can be a **[PjTimescaledUnit](Project.PjTimescaleUnit.md)** constant. The default value is **pjTimescaleDays**.|
| _TimescaleUnitCount_|Optional|**Long**|Specifies the number of timescale units to be included in the update. The default value is 1. For example, if the unit is **pjTimescaleWeeks**, a value of 5 indicates five weeks.|
| _StartDate_|Optional|**Variant**|Specifies the start date for the update.|
| _FinishDate_|Optional|**Variant**|Specifies the finish date for the update.|


## Return value

**Nothing**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]