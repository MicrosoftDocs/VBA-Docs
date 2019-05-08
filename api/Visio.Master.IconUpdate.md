---
title: Master.IconUpdate property (Visio)
keywords: vis_sdr.chm10713670
f1_keywords:
- vis_sdr.chm10713670
ms.prod: visio
api_name:
- Visio.Master.IconUpdate
ms.assetid: 3978c650-47d5-e961-53c2-d99dd4c2ca7c
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.IconUpdate property (Visio)

Determines whether a master icon is updated manually or automatically. Read/write.


## Syntax

_expression_. `IconUpdate`

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Return value

Integer


## Remarks

The following constants declared by the Visio type library in  **VisMasterProperties** show the possible values for the **IconUpdate** property.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visManual**|0 |Update icon manually.|
| **visAutomatic**|1 |Update icon automatically from shape geometry data.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]