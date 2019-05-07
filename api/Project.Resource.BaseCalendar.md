---
title: Resource.BaseCalendar property (Project)
ms.prod: project-server
api_name:
- Project.Resource.BaseCalendar
ms.assetid: f6893deb-6faa-2d36-6633-5186f2af5765
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.BaseCalendar property (Project)

Gets or sets the name of the base calendar for a resource. Read/write  **String**.


## Syntax

_expression_. `BaseCalendar`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Remarks

The  **BaseCalendar** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]