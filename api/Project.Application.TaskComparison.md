---
title: Application.TaskComparison method (Project)
keywords: vbapj.chm2184
f1_keywords:
- vbapj.chm2184
ms.prod: project-server
api_name:
- Project.Application.TaskComparison
ms.assetid: 61d0c322-39a3-f731-3662-f6cf6709bb12
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.TaskComparison method (Project)

In a project comparison report, shows the Gantt Chart view in all three project plans, to compare tasks.


## Syntax

_expression_. `TaskComparison`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

Use the  **CompareProjectVersions** method to create a project comparison report, or choose **Compare Projects** on the **PROJECT** ribbon.



After you run a  **Compare Projects** command, Project displays the **COMPARE PROJECTS** ribbon. The **TaskComparison** method is equivalent to the **Task Comparison** command on the **COMPARE PROJECTS** ribbon.

To compare resources in a comparison report, use the  **[ResourceComparison](Project.Application.ResourceComparison.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]