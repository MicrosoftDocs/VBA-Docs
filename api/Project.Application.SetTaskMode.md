---
title: Application.SetTaskMode method (Project)
keywords: vbapj.chm90
f1_keywords:
- vbapj.chm90
ms.prod: project-server
api_name:
- Project.Application.SetTaskMode
ms.assetid: 0d800877-9cd9-97e0-6912-6a8d5f596276
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SetTaskMode method (Project)

Changes the mode of the selected tasks, to manually scheduled or automatically scheduled.


## Syntax

_expression_. `SetTaskMode`( `_Manual_`, `_IsStickyDates_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Manual_|Optional|**Boolean**|If  **True**, changes the selected tasks to manually scheduled. If **False**, changes the tasks to automatically scheduled.|
| _IsStickyDates_|Optional|**Boolean**|If  **True**, when a manually scheduled task is changed to automatically scheduled, the constraint type is set to **Start No Earlier Than** and the constraint date is set to the previous start date.|

## Return value

 **Boolean**


## Remarks

The  **SetTaskMode** method corresponds to the **Manually Schedule** command and the **Auto Schedule** command on the **TASK** ribbon.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]