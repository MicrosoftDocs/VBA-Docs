---
title: Application.AutoFilter method (Project)
keywords: vbapj.chm22
f1_keywords:
- vbapj.chm22
ms.prod: project-server
api_name:
- Project.Application.AutoFilter
ms.assetid: 391d5a61-cba3-9e28-c448-d0befcc456c7
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AutoFilter method (Project)

Activates or deactivates the AutoFilter feature for the active project.


## Syntax

_expression_. `AutoFilter`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The  **AutoFilter** method toggles the AutoFilter feature on and off. It has the same effect as the **AutoFilter** command on the filter drop-down menu on the **View** tab for **Gantt Chart Tools** in the Ribbon. If column headings show the AutoFilter drop-down menu, executing the AutoFilter method hides the AutoFilter menus for columns in all sheet views in the active project.

To set an AutoFilter, see the  **[SetAutoFilter](Project.Application.SetAutoFilter.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]