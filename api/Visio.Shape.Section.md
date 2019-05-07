---
title: Shape.Section property (Visio)
keywords: vis_sdr.chm11214300
f1_keywords:
- vis_sdr.chm11214300
ms.prod: visio
api_name:
- Visio.Shape.Section
ms.assetid: e87823aa-fd7c-e222-417b-a167d2e0898a
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Section property (Visio)

Returns the requested  **Section** object belonging to a shape. Read-only.


## Syntax

_expression_.**Section** (_Index_)

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Integer**|A section index.|

## Return value

Section


## Remarks

Constants that represent sections are prefixed with  **visSection** and are declared by the Microsoft Visio type library in **[VisSectionIndices](Visio.vissectionindices.md)**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]