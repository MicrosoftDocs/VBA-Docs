---
title: Cell.CellColor property (Project)
ms.prod: project-server
api_name:
- Project.Cell.CellColor
ms.assetid: 30d67933-a9ce-9e57-f7ac-c4af2f485959
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.CellColor property (Project)

Gets or sets the color of the cell background. Read/write  **PjColor**.


## Syntax

_expression_. `CellColor`

_expression_ A variable that represents a [Cell](./Project.Cell.md) object.


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

To use a hexadecimal RGB value for the cell color, see the **[CellColorEx](Project.Cell.CellColorEx.md)** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]