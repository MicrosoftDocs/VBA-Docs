---
title: Resource.OvertimeWork property (Project)
ms.prod: project-server
api_name:
- Project.Resource.OvertimeWork
ms.assetid: c9656656-2e8f-d09d-8c91-ebf4d42ccaba
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.OvertimeWork property (Project)

Gets the overtime work for a resource. Read-only  **Variant**.


## Syntax

_expression_. `OvertimeWork`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Remarks

The **OvertimeWork** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]