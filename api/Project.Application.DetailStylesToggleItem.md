---
title: Application.DetailStylesToggleItem method (Project)
keywords: vbapj.chm960
f1_keywords:
- vbapj.chm960
ms.prod: project-server
api_name:
- Project.Application.DetailStylesToggleItem
ms.assetid: 744022ac-e5c1-ee5a-c02b-c6962c821c55
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DetailStylesToggleItem method (Project)

Toggles the display of a timescale data field in a usage view.


## Syntax

_expression_. `DetailStylesToggleItem`( `_Item_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Optional|**Long**|The timescale data field to show or remove. The default value is  **pjWork**.|

## Return value

 **Boolean**


## Remarks

If the active view is the  **Resource Usage** view, can be one of the following **PjTimescaledData** constants:


|||
|:-----|:-----|
|**pjActualCost**|**pjCumulativeCost**|
|**pjActualOvertimeWork**|**pjCumulativeWork**|
|**pjActualWork**|**pjCV**|
|**pjACWP**|**pjOverallocation**|
|**pjAllAssignmentRows**|**pjOvertimeWork**|
|**pjAllResourceRows**|**pjPeakUnits**|
|**pjBaselineCost**|**pjPercentAllocation**|
|**pjBaselineWork**|**pjRegularWork**|
|**pjBaseline1-10Cost**|**pjRemainingAvailability**|
|**pjBaseline1-10Work**|**pjSV**|
|**pjBCWP**|**pjWork**|
|**pjBCWS**|**pjWorkAvailability**|
|**pjCost**||

If the active view is the  **Task Usage** view, can be one of the following **PjTimescaledData** constants:


|||
|:-----|:-----|
|**pjActualCost**|**pjCumulativeCost**|
|**pjActualFixedCost**|**pjCumulativeWork**|
|**pjActualOvertimeWork**|**pjCV**|
|**pjActualWork**|**pjCVP**|
|**pjACWP**|**pjFixedCost**|
|**pjAllAssignmentRows**|**pjOverallocation**|
|**pjAllTaskRows**|**pjOvertimeWork**|
|**pjBaselineCost**|**pjPeakUnits**|
|**pjBaselineWork**|**pjPercentAllocation**|
|**pjBaseline1-10Cost**|**pjPctComplete**|
|**pjBaseline1-10Work**|**pjRegularWork**|
|**pjBCWP**|**pjSPIT**|
|**pjBCWS**|**pjSV**|
|**pjCost**|**pjSVP**|
|**pjCPI**|**pjWork**|
|**pjCumPctComplete**||

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]