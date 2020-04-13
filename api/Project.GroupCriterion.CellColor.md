---
title: GroupCriterion.CellColor property (Project)
ms.prod: project-server
api_name:
- Project.GroupCriterion.CellColor
ms.assetid: dcddcac1-e935-9e60-9611-5bf77267c5f1
ms.date: 06/08/2017
localization_priority: Normal
---


# GroupCriterion.CellColor property (Project)

Gets or sets the color of the cell background for a field used as a criterion in a group definition. Read/write  **PjColor**.


## Syntax

_expression_. `CellColor`

_expression_ A variable that represents a [GroupCriterion](./Project.GroupCriterion.md) object.


## Remarks

The **CellColor** property can be one of the following **[PjColor](Project.PjColor.md)** constants:


|||
|:-----|:-----|
|**pjColorAutomatic**|**pjNavy**|
|**pjAqua**|**pjOlive**|
|**pjBlack**|**pjPurple**|
|**pjBlue**|**pjRed**|
|**pjFuchsia**|**pjSilver**|
|**pjGray**|**pjTeal**|
|**pjGreen**|**pjYellow**|
|**pjLime**|**pjWhite**|
|**pjMaroon**||

To use a hexadecimal RGB value for the cell color, see the **[CellColorEx](Project.GroupCriterion2.CellColorEx.md)** property of the **GroupCriterion2** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]