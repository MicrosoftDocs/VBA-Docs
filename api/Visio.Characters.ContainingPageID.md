---
title: Characters.ContainingPageID property (Visio)
keywords: vis_sdr.chm10251930
f1_keywords:
- vis_sdr.chm10251930
ms.prod: visio
api_name:
- Visio.Characters.ContainingPageID
ms.assetid: 095cd4fc-1aa1-338a-eb9a-dedb63c2c1ad
ms.date: 06/08/2017
localization_priority: Normal
---


# Characters.ContainingPageID property (Visio)

Returns the ID of the page that contains an object. Read-only.


## Syntax

_expression_.**ContainingPageID**

_expression_ A variable that represents a **[Characters](Visio.Characters.md)** object.


## Return value

Long


## Remarks

If the object is not in a  **Page** object, the **ContainingPageID** property returns -1. For example, if a **Shape** object belongs to a **Masters** collection, the **ContainingPageID** property returns -1.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]