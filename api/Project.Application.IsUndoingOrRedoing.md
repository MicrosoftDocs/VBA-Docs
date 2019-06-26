---
title: Application.IsUndoingOrRedoing method (Project)
ms.prod: project-server
api_name:
- Project.Application.IsUndoingOrRedoing
ms.assetid: e0e5ddc7-aa22-0d43-1de6-83a260d57608
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.IsUndoingOrRedoing method (Project)

Indicates whether Project is currently executing an undo or redo action.


## Syntax

_expression_. `IsUndoingOrRedoing`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

 Use the **[Application.OnUndoOrRedo](Project.Application.OnUndoOrRedo.md)** event to listen for specific undo or redo actions.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]