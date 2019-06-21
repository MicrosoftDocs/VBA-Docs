---
title: Selection.ContainingPage property (Visio)
keywords: vis_sdr.chm11113305
f1_keywords:
- vis_sdr.chm11113305
ms.prod: visio
api_name:
- Visio.Selection.ContainingPage
ms.assetid: dca54861-d6c6-9d39-2a49-2070a578607f
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ContainingPage property (Visio)

Returns the page that contains an object.


## Syntax

_expression_. `ContainingPage`

_expression_ A variable that represents a **[Selection](Visio.Selection.md)** object.


## Return value

Page


## Remarks

If the object isn't in a  **Page** object, the **ContainingPage** property returns **Nothing**. For example, if a **Shape** object belongs to a **Masters** collection, the **ContainingPage** property returns **Nothing**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]