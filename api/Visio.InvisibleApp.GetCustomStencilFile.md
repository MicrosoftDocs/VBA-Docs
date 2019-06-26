---
title: InvisibleApp.GetCustomStencilFile method (Visio)
keywords: vis_sdr.chm17562115
f1_keywords:
- vis_sdr.chm17562115
ms.prod: visio
api_name:
- Visio.InvisibleApp.GetCustomStencilFile
ms.assetid: 8ccb6786-de34-5fc2-83ed-aae5f9f7a191
ms.date: 06/26/2019
localization_priority: Normal
---


# InvisibleApp.GetCustomStencilFile method (Visio)

Returns the path to the specified custom stencil used to populate certain galleries in the Microsoft Visio user interface.


## Syntax

_expression_.**GetCustomStencilFile** (_StencilType_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _StencilType_|Required| **[VisBuiltInStencilTypes](Visio.VisBuiltInStencilTypes.md)**|The stencil to retrieve. Must be one of the **VisBuiltInStencilTypes** constants.|

## Return value

**String**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]