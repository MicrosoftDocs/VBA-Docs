---
title: Application.ViewShowUnitAvailability method (Project)
keywords: vbapj.chm931
f1_keywords:
- vbapj.chm931
ms.prod: project-server
api_name:
- Project.Application.ViewShowUnitAvailability
ms.assetid: 900af4b4-dd2d-483e-b207-6d199c51092b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ViewShowUnitAvailability method (Project)

Displays unit availability information in the active Resource Graph view.


## Syntax

_expression_. `ViewShowUnitAvailability`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Remarks

The  **ViewShowUnitAvailability** method has no effect unless the active window contains the Resource Graph view. The **ViewShowUnitAvailability** method is not available for material resources and returns a trappable error (error code 1100) when applied to material resources.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]