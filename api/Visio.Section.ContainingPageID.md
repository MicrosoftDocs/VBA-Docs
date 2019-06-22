---
title: Section.ContainingPageID property (Visio)
keywords: vis_sdr.chm15751695
f1_keywords:
- vis_sdr.chm15751695
ms.prod: visio
api_name:
- Visio.Section.ContainingPageID
ms.assetid: 9c32b32a-7052-be2c-ee2a-fc145be626eb
ms.date: 06/08/2017
localization_priority: Normal
---


# Section.ContainingPageID property (Visio)

Returns the ID of the page that contains the section. Read-only. .


## Syntax

_expression_. `ContainingPageID`

_expression_ A variable that represents a **[Section](Visio.Section.md)** object.


## Return value

Long


## Remarks

If the section is not in a  **Page** object, the **ContainingPageID** property returns -1.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]