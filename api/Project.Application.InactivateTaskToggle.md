---
title: Application.InactivateTaskToggle method (Project)
keywords: vbapj.chm91
f1_keywords:
- vbapj.chm91
ms.prod: project-server
api_name:
- Project.Application.InactivateTaskToggle
ms.assetid: af937c95-b434-95b8-7ea4-848c25ca30bc
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.InactivateTaskToggle method (Project)

Toggles the state of a task between inactive and active.


## Syntax

_expression_. `InactivateTaskToggle`( `_MakeActive_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MakeActive_|Optional|**Variant**|Specifies whether to make the task active. If the task is active,  **True** has no effect. If the task is inactive, **True** makes the task active.|

## Return value

 **Boolean**


## Remarks

The **InactivateTaskToggle** method corresponds to the **Inactivate** command in the **Tasks** group of the **Task** tab on the Ribbon.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]