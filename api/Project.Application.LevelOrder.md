---
title: Application.LevelOrder property (Project)
ms.prod: project-server
api_name:
- Project.Application.LevelOrder
ms.assetid: c8cf70bb-7808-48c4-43b4-c7f693d4613d
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.LevelOrder property (Project)

Gets or sets the order in which tasks with overallocations will be delayed when resources are leveled. Read/write  **PjLevelOrder**.


## Syntax

_expression_. `LevelOrder`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

The **LevelOrder** property can be one of the following **[PjLevelOrder](Project.PjLevelOrder.md)** constants: **pjLevelID**, **pjLevelStandard**, or **pjLevelPriority**.

You can also set the **LevelOrder** property in the **Resource Leveling** dialog box. To access the setting, click **Leveling Options** on the **Resource** tab of the Ribbon, and then set the **Leveling order** drop-down list.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]