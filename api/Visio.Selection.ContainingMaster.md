---
title: Selection.ContainingMaster property (Visio)
keywords: vis_sdr.chm11113300
f1_keywords:
- vis_sdr.chm11113300
ms.prod: visio
api_name:
- Visio.Selection.ContainingMaster
ms.assetid: 9eae609f-2d55-2180-ea9b-cf1f8ec7b7b3
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.ContainingMaster property (Visio)

Returns the  **Master** object that contains an object. Read-only.


## Syntax

_expression_. `ContainingMaster`

_expression_ A variable that represents a **[Selection](Visio.Selection.md)** object.


## Return value

Master


## Remarks

If the object isn't in a  **Master** object, the **ContainingMaster** property returns **Nothing**. For example, if a **Shape** object belongs to the **Shapes** collection of a **Page** object, the **ContainingMaster** property returns **Nothing**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]