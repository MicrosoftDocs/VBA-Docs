---
title: Application.LevelProposedBookings property (Project)
ms.prod: project-server
api_name:
- Project.Application.LevelProposedBookings
ms.assetid: 34b1d355-a5c5-38c2-9502-064ecd81906e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.LevelProposedBookings property (Project)

Gets or sets a value that indicates whether proposed assignment bookings will be leveled. Read/write  **Boolean**.


## Syntax

_expression_. `LevelProposedBookings`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

When the booking type is proposed, the resource is not yet assigned to the project and the resource calendar remains unchanged by the proposed assignment.

The **[LevelingOptions](Project.Application.LevelingOptions.md)** method includes a LevelProposedBookings parameter.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]