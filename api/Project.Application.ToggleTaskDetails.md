---
title: Application.ToggleTaskDetails method (Project)
keywords: vbapj.chm2298
f1_keywords:
- vbapj.chm2298
ms.prod: project-server
api_name:
- Project.Application.ToggleTaskDetails
ms.assetid: c27dffe7-6814-85f5-9c49-21e0efb12cd1
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ToggleTaskDetails method (Project)

Shows or hides the **Task Form** in the bottom pane of a split view.


## Syntax

_expression_. `ToggleTaskDetails`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The **ToggleTaskDetails** method corresponds to selecting or clearing the **Details** check box in the **Split View** group on the **View** tab under **Gantt Chart Tools** on the ribbon, where **Task Form** is selected in the **Details** drop-down list.

You can use  **ToggleTaskDetails** to add a **Task Form** split view to other views except an empty Timeline view.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]