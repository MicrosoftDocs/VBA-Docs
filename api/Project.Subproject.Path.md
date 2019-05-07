---
title: Subproject.Path property (Project)
ms.prod: project-server
api_name:
- Project.Subproject.Path
ms.assetid: 57bd6c44-5a2e-a2c8-c733-4c46e32be780
ms.date: 06/08/2017
localization_priority: Normal
---


# Subproject.Path property (Project)

Gets or sets the path to the source project. Read/write  **String**.


## Syntax

_expression_.**Path**

_expression_ A variable that represents a [Subproject](./Project.Subproject.md) object.


## Remarks

The  **Path** property (**Subproject** object) can be set only if the **[LinkToSource](Project.Subproject.LinkToSource.md)** property for the subproject has been set to **True**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]