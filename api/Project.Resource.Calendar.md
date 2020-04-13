---
title: Resource.Calendar property (Project)
ms.prod: project-server
api_name:
- Project.Resource.Calendar
ms.assetid: 93bf12ea-ba8e-3b98-cc28-7af5168b514f
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.Calendar property (Project)

Gets a **[Calendar](Project.Calendar.md)** object representing a calendar for the resource. Read-only **Calendar**.


## Syntax

_expression_. `Calendar`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Remarks

The **Calendar** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.

For an example that resets the project calendar, see the **[Calendar](Project.Project.Calendar.md)** property of the **Project** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]