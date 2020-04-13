---
title: Application.TogglePreventResOveralloc method (Project)
keywords: vbapj.chm1501
f1_keywords:
- vbapj.chm1501
ms.prod: project-server
api_name:
- Project.Application.TogglePreventResOveralloc
ms.assetid: 7b6686ab-58c6-e1de-cbb1-618495d5c8ba
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.TogglePreventResOveralloc method (Project)

Toggles the **Prevent Overallocations** command for the Team Planner view.


## Syntax

_expression_. `TogglePreventResOveralloc`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

When the **Prevent Overallocations** command is active, Project automatically moves tasks so that resources do not become overallocated because of changes made in the Team Planner view. Overallocations that exist when the **Prevent Overallocations** command is made active are also resolved.

The **TogglePreventResOveralloc** method corresponds to the **Prevent Overallocations** command on the **Format** tab under **Team Planner Tools** on the ribbon.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]