---
title: Application.PMText property (Project)
ms.prod: project-server
api_name:
- Project.Application.PMText
ms.assetid: a52193c7-2a74-c3b8-357b-ea7637309d14
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PMText property (Project)

Gets the text that Project displays next to evening hours in the 12-hour time format. Read-only  **String**.


## Syntax

_expression_. `PMText`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

Project sets the **AMText** and **PMText** properties equal to the corresponding values in the **Regional and Language Options** dialog box opened from the Microsoft Windows Control Panel.


> [!NOTE] 
> Although the VBA Object Browser shows  **PMText** as read-write, you cannot set the value using the **PMText** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]