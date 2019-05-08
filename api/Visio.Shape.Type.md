---
title: Shape.Type property (Visio)
keywords: vis_sdr.chm11214595
f1_keywords:
- vis_sdr.chm11214595
ms.prod: visio
api_name:
- Visio.Shape.Type
ms.assetid: 0d7438d2-e2df-2045-1a2f-608eca530bc1
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Type property (Visio)

Returns the type of the object. Read-only.


## Syntax

_expression_.**Type**

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Return value

Integer


## Remarks

Type value constants for  **Shape** objects (the possible values that the **Type** property of a **Shape** object returns) are declared by the Visio type library in **[VisShapeTypes](Visio.visshapetypes.md)**.

If a  **Shape** object is type **visTypeForeignObject**, use the **ForeignType** property to determine the type of foreign object represented by the object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]