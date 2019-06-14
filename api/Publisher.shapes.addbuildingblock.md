---
title: Shapes.AddBuildingBlock method (Publisher)
keywords: vbapb10.chm2162768
f1_keywords:
- vbapb10.chm2162768
ms.prod: publisher
ms.assetid: d875e97e-3519-4a88-916d-ec1a32654581
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddBuildingBlock method (Publisher)

Adds a **[BuildingBlock](Publisher.BuildingBlock.md)** object and returns a **[Shape](Publisher.Shape.md)** object on the page that represents the building block.


## Syntax

_expression_.**AddBuildingBlock** (_BBlockIn_, _Left_, _Top_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_BBlockIn_|Required| **BuildingBlock**|The building block to return as a shape.|
|_Left_|Required| **Variant**|The position of the left edge of the shape that represents the building block.|
|_Top_|Required| **Variant**|The position of the top edge of the shape that represents the building block.|


## Return value

Shape



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]