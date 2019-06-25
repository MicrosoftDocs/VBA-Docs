---
title: Application.GetCustomStencilFile method (Visio)
keywords: vis_sdr.chm10062115
f1_keywords:
- vis_sdr.chm10062115
ms.prod: visio
api_name:
- Visio.Application.GetCustomStencilFile
ms.assetid: 10c8ec1d-f4e0-07dd-4487-40f85cbf5497
ms.date: 06/26/2019
localization_priority: Normal
---


# Application.GetCustomStencilFile method (Visio)

Returns the path to the specified custom stencil used to populate certain galleries in the Microsoft Visio user interface.


## Syntax

_expression_.**GetCustomStencilFile** (_StencilType_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _StencilType_|Required| **[VisBuiltInStencilTypes](Visio.VisBuiltInStencilTypes.md)**|The stencil to retrieve. Must be one of the **VisBuiltInStencilTypes** constants.|

## Return value

**String**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]