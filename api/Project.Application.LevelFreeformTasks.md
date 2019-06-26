---
title: Application.LevelFreeformTasks property (Project)
ms.prod: project-server
api_name:
- Project.Application.LevelFreeformTasks
ms.assetid: d9a9abca-0efa-ea38-3665-7f7b7ecccc9e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.LevelFreeformTasks property (Project)

 **True** if manually scheduled tasks are leveled; otherwise, **False**. Read/write **Boolean**.


## Syntax

_expression_. `LevelFreeformTasks`

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Remarks

The  **LevelFreeformTasks** property corresponds to the **Level manually scheduled tasks** check box in the **Resource Leveling** dialog box. To access the **Resource Leveling** dialog box, on the **Resource** tab of the Ribbon, choose **Leveling Options**.

To set other leveling options, see the  **[LevelingOptionsEx](Project.Application.LevelingOptionsEx.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]