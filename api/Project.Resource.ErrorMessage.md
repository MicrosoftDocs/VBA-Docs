---
title: Resource.ErrorMessage property (Project)
ms.prod: project-server
api_name:
- Project.Resource.ErrorMessage
ms.assetid: cb78732f-8c9c-df97-b6bc-c4eb52f4bf16
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.ErrorMessage property (Project)

Gets errors reported by the **Import Resources Wizard** and by local resource error checks. Read-only **String**.


## Syntax

_expression_. `ErrorMessage`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Remarks

The **ErrorMessage** property is used by the **Import Resources Wizard** while saving the enterprise resource pool and when **[CheckResourceErrors](Project.Application.CheckResourceErrors.md)** and **[EnterpriseResourcesImport](Project.Application.EnterpriseResourcesImportEx.md)** methods are used.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]