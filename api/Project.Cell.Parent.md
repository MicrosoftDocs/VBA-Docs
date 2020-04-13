---
title: Cell.Parent property (Project)
ms.prod: project-server
api_name:
- Project.Cell.Parent
ms.assetid: 8e2f9a5d-b914-f9e1-b922-ade8fb7ade01
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Parent property (Project)

Gets the parent of the **Cell** object. Read-only **Object**.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a [Cell](./Project.Cell.md) object.


## Remarks

The parent of a **Cell** object can be the **Application** or a **Project**. For example, the statement `Application.ActiveCell.Parent` gets the name of the active project.

Use the **Parent** property to access the properties or methods of the parent of an object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]