---
title: Project.DefaultFixedCostAccrual property (Project)
ms.prod: project-server
api_name:
- Project.Project.DefaultFixedCostAccrual
ms.assetid: 24acadcb-6eed-6b5e-ca50-5b509a7e4af0
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.DefaultFixedCostAccrual property (Project)

Gets or sets the default method used to accrue fixed task costs in the project. Read/write  **PjAccrueAt**.


## Syntax

_expression_. `DefaultFixedCostAccrual`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

The **DefaultFixedCostAccrual** property can be one of the following **[PjAccrueAt](Project.PjAccrueAt.md)** constants: **pjStart**, **pjEnd**, or **pjProrated**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]