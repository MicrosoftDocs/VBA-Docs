---
title: Application.TwelveHourTimeFormat property (Project)
ms.prod: project-server
api_name:
- Project.Application.TwelveHourTimeFormat
ms.assetid: 899caa96-da4e-8ee6-988a-6cef64a1a46c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.TwelveHourTimeFormat property (Project)

 **True** if Project returns the time in a 12-hour format. **False** if the time is returned in a 24-hour format. Read-only **Boolean**.


## Syntax

_expression_. `TwelveHourTimeFormat`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

Project internally sets the **TwelveHourTimeFormat** property equal to the corresponding value in the **Regional and Language Options** dialog box of the Microsoft Windows Control Panel.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]