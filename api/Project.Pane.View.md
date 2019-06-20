---
title: Pane.View method (Project)
ms.prod: project-server
api_name:
- Project.Pane.View
ms.assetid: a29aa7d4-e712-bbf4-96dd-e0fdeab70ba2
ms.date: 06/08/2017
localization_priority: Normal
---


# Pane.View method (Project)

Returns the active  **View** object.


## Syntax

_expression_.**View**

_expression_ A variable that represents a [Pane](./Project.Pane.md) object.


## Return value

 **View**


## Example

The following statement prints the name of the view in the Immediate window in the VBE. For example, if the Team Planner view is active, the statement prints "Team Planner".


```vb
Debug.Print ActiveWindow.ActivePane.View.Name
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]