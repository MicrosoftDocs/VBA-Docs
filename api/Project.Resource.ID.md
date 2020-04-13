---
title: Resource.ID property (Project)
ms.prod: project-server
api_name:
- Project.Resource.ID
ms.assetid: 15e18fda-ca6d-c81b-55c8-ad21605f75fc
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.ID property (Project)

Gets the identification number of a resource. Read-only  **Long**.


## Syntax

_expression_.**ID**

 _expression_ An expression that returns a [Resource](./Project.Resource.md) object.


## Remarks

The **ID** property changes when a resource moves to a new location in a view such as the **Resource Sheet**. Use the **UniqueID** property if you want a constant reference to a resource.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]