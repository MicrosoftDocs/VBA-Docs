---
title: Project.Container property (Project)
ms.prod: project-server
api_name:
- Project.Project.Container
ms.assetid: 34969587-b74d-3425-0f4f-af7d90221b10
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Container property (Project)

Gets the object that contains the embedded project. Read-only  **Object**.


## Syntax

_expression_. `Container`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

Use the **Container** property to access the properties or methods of the object that contains the project. If the container doesn't support automation or the project is not embedded, the **Container** property fails with run-time error 1004.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]