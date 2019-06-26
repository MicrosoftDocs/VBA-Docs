---
title: Application.Visible property (Project)
ms.prod: project-server
api_name:
- Project.Application.Visible
ms.assetid: 43bf25de-4908-1fad-e5d5-9fba21e8b03c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Visible property (Project)

 **True** if the application is visible. Read/write **Boolean**.


## Syntax

_expression_.**Visible**

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

The  **Visible** property can only be set to **False** if the **Application**. **[UserControl](Project.Application.UserControl.md)** property is **False** and there are no visible projects. If the **UserControl** property is **True**, the Project application is under user control rather than programmatic control, and the **Visible** property is also **True**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]