---
title: Application.SetShowTaskSuggestions method (Project)
keywords: vbapj.chm2177
f1_keywords:
- vbapj.chm2177
ms.prod: project-server
api_name:
- Project.Application.SetShowTaskSuggestions
ms.assetid: 650dd088-9b38-8706-900d-dad7a6ebf4fd
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SetShowTaskSuggestions method (Project)

Sets the global  **Show Suggestions** option for tasks.


## Syntax

_expression_. `SetShowTaskSuggestions`( `_Set_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Set_|Optional|**Variant**|If  **True**, turns on the **Show Suggestions** option. The default value is **False**.|

## Return value

 **Boolean**


## Remarks

The **Show Suggestions** option is in the drop-down **Inspect Task** menu on the **Task** tab of the ribbon. You can override the global setting for a specific task by selecting or clearing the **Show warning and suggestion indicators for this task** check box in the **Task Inspector** pane.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]