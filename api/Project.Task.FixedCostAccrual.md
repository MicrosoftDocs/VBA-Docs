---
title: Task.FixedCostAccrual property (Project)
ms.prod: project-server
api_name:
- Project.Task.FixedCostAccrual
ms.assetid: 22a76efc-de26-3687-6ffe-674478c48767
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.FixedCostAccrual property (Project)

Gets or sets the way the task accrues fixed costs. Read/write  **PjAccrueAt**.


## Syntax

_expression_. `FixedCostAccrual`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

The  **FixedCostAccrual** property can be one of the following **[PjAccrueAt](Project.PjAccrueAt.md)** constants: **pjEnd**, **pjProrated**, or **pjStart**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]