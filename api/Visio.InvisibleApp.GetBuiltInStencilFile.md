---
title: InvisibleApp.GetBuiltInStencilFile method (Visio)
keywords: vis_sdr.chm17562110
f1_keywords:
- vis_sdr.chm17562110
ms.prod: visio
api_name:
- Visio.InvisibleApp.GetBuiltInStencilFile
ms.assetid: 2f8e28a9-67bd-31fd-25f1-f684dfeeeca8
ms.date: 06/26/2019
localization_priority: Normal
---


# InvisibleApp.GetBuiltInStencilFile method (Visio)

Returns the file path to the specified built-in, hidden stencil used to populate certain galleries in the Microsoft Visio user interface.


## Syntax

_expression_.**GetBuiltInStencilFile** (_StencilType_, _MeasurementSystem_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _StencilType_|Required| **[VisBuiltInStencilTypes](Visio.VisBuiltInStencilTypes.md)**|The stencil to retrieve. Must be one of the **VisBuiltInStencilTypes** constants.|
| _MeasurementSystem_|Required| **[VisMeasurementSystem](Visio.vismeasurementsystem.md)**|The measurement system for the stencil.|

## Return value

**String**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]