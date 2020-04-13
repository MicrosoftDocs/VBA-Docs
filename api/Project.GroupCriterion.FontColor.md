---
title: GroupCriterion.FontColor property (Project)
ms.prod: project-server
api_name:
- Project.GroupCriterion.FontColor
ms.assetid: 9765d7a2-0f6e-8fa1-210a-9ad138bae9a7
ms.date: 06/08/2017
localization_priority: Normal
---


# GroupCriterion.FontColor property (Project)

Gets or sets the color of the font for a field used as a criterion in a group definition. Read/write  **PjColor**.


## Syntax

_expression_. `FontColor`

_expression_ A variable that represents a [GroupCriterion](./Project.GroupCriterion.md) object.


## Remarks

The **FontColor** property can be one of the following **[PjColor](Project.PjColor.md)** constants:


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

To use a hexadecimal RGB value, see the **[FontColorEx](Project.GroupCriterion2.FontColorEx.md)** property of the **GroupCriterion2** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]