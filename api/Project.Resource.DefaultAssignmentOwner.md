---
title: Resource.DefaultAssignmentOwner property (Project)
ms.prod: project-server
api_name:
- Project.Resource.DefaultAssignmentOwner
ms.assetid: 41f08732-0f5a-e366-dbc0-54aab1a89fe2
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.DefaultAssignmentOwner property (Project)

Sets or gets the user name responsible for providing progress updates for assignments made to the resource. Read/write  **String**.


## Syntax

_expression_. `DefaultAssignmentOwner`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Remarks

The **DefaultAssignmentOwner** property affects all assignments for the resource. The property must be set to a valid Project Server user or **null**.

The **DefaultAssignmentOwner** property is available only in Project Professional.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]