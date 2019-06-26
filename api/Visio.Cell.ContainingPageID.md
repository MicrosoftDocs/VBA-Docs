---
title: Cell.ContainingPageID property (Visio)
keywords: vis_sdr.chm10151695
f1_keywords:
- vis_sdr.chm10151695
ms.prod: visio
api_name:
- Visio.Cell.ContainingPageID
ms.assetid: 0d4c97cc-d84e-c13e-759b-8805114d191e
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.ContainingPageID property (Visio)

Returns the ID of the page that contains an object. Read-only.


## Syntax

_expression_.**ContainingPageID**

_expression_ A variable that represents a **[Cell](Visio.Cell.md)** object.


## Return value

Long


## Remarks

If the object is not in a  **Page** object, the **ContainingPageID** property returns -1. For example, if a **Shape** object belongs to a **Masters** collection, the **ContainingPageID** property returns -1.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]