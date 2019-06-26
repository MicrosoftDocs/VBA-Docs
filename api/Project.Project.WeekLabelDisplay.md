---
title: Project.WeekLabelDisplay property (Project)
keywords: vbapj.chm132820
f1_keywords:
- vbapj.chm132820
ms.prod: project-server
api_name:
- Project.Project.WeekLabelDisplay
ms.assetid: d21cd816-06a3-89b0-b56a-9c1b56151209
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.WeekLabelDisplay property (Project)

Gets or sets the abbreviation for "week" that is displayed for values such as durations, delays, slack, and work. Read/write  **Integer**.


## Syntax

_expression_. `WeekLabelDisplay`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

The  **WeekLabelDisplay** property corresponds to the **Weeks** list on the **Advanced** tab of the **Project Options** dialog box. For example, setting the **WeekLabelDisplay** property to 1 sets the **Weeks** list to the second value in the list ("wk").

Values of the  **WeekLabelDisplay** property can be 0 to 2.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]