---
title: Assignment.FixedMaterialAssignment property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.FixedMaterialAssignment
ms.assetid: 16593466-1d5e-27b3-110d-e5cfeb165355
ms.date: 06/08/2017
localization_priority: Normal
---


# Assignment.FixedMaterialAssignment property (Project)

 **True** if the consumption of the assigned material resource occurs in a single, fixed amount. **False** if the consumption occurs at an hourly rate. Read-only **Boolean**.


## Syntax

_expression_. `FixedMaterialAssignment`

_expression_ A variable that represents an [Assignment](./Project.Assignment.md) object.


## Remarks

The **FixedMaterialAssignment** property returns **True** if the assignment is for a work (non-material) resource.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]