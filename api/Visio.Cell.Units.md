---
title: Cell.Units property (Visio)
keywords: vis_sdr.chm10114620
f1_keywords:
- vis_sdr.chm10114620
ms.prod: visio
api_name:
- Visio.Cell.Units
ms.assetid: 075cfda9-8b7a-550b-cf72-b8044c3d461a
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Units property (Visio)

Indicates the unit of measure associated with a  **Cell** object. Read-only.


## Syntax

_expression_.**Units**

_expression_ A variable that represents a **[Cell](Visio.Cell.md)** object.


## Return value

Integer


## Remarks

The  **Units** property can be used to determine the unit of measure currently associated with a cell's value. The various unit codes are declared by the Visio type library in member **[VisUnitCodes](Visio.visunitcodes.md)**. For example, a cell's width might be expressed in inches (**visInches**) or in centimeters (**visCentimeters**). In some cases, a program might behave differently depending on whether a cell's value is in metric or in imperial units.

For a list of valid unit codes, see [About units of measure](../visio/Concepts/about-units-of-measure-visio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]