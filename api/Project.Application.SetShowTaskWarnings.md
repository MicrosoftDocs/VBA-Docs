---
title: Application.SetShowTaskWarnings method (Project)
keywords: vbapj.chm2176
f1_keywords:
- vbapj.chm2176
ms.prod: project-server
api_name:
- Project.Application.SetShowTaskWarnings
ms.assetid: 43ccb666-c61d-e26a-2645-9fa2cb4b3d72
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SetShowTaskWarnings method (Project)

Sets the global  **Show Warnings** option for tasks.


## Syntax

_expression_. `SetShowTaskWarnings`( `_Set_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Set_|Optional|**Variant**|If  **True**, turns on the **Show Warnings** option. The default value is **False**.|

## Return value

 **Boolean**


## Remarks

The **Show Warnings** option is in the drop-down **Inspect Task** menu on the **TASK** ribbon. You can override the global setting for a specific task by selecting or clearing the **Show warning and suggestion indicators for this task** check box in the **Task Inspector** pane.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]