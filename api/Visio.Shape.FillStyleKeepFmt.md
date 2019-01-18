---
title: Shape.FillStyleKeepFmt Property (Visio)
keywords: vis_sdr.chm11213530
f1_keywords:
- vis_sdr.chm11213530
ms.prod: visio
api_name:
- Visio.Shape.FillStyleKeepFmt
ms.assetid: 39fc0329-322e-fd96-2c42-43bdcd170c02
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.FillStyleKeepFmt Property (Visio)

Applies a fill style to an object while preserving local formatting. Read/write.


## Syntax

 _expression_. `FillStyleKeepFmt`

 _expression_ A variable that represents a [Shape](./Visio.Shape.md) object.


## Return value

String


## Remarks

Setting a style to a nonexistent style generates an error. Setting one type of style to another type (for example, setting the  **FillStyleKeepFmt** property to a line style) does nothing. Setting one type of style to another type that has more than one set of attributes changes only the appropriate attributes (for example, setting the **FillStyleKeepFmt** property to a style that has line, text, and fill attributes changes only the fill attributes).

Beginning with Microsoft Visio 2002, setting  **FillStyleKeepFmt** to an empty string ("") causes the master's style to be reapplied to the selection or shape. (Earlier versions generate a "no such style" exception.) If the selection or shape has no master, its style remains unchanged.


