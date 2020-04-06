---
title: Master.AddGuide method (Visio)
keywords: vis_sdr.chm10716035
f1_keywords:
- vis_sdr.chm10716035
ms.prod: visio
api_name:
- Visio.Master.AddGuide
ms.assetid: 7beba614-244b-f559-50c7-5156ca4510b1
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.AddGuide method (Visio)

Adds a guide to a master.


## Syntax

_expression_. `AddGuide`( `_Type_` , `_xPos_` , `_yPos_` )

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **Integer**|The type of guide to add.|
| _xPos_|Required| **Double**|The x-coordinate of a point on the guide.|
| _yPos_|Required| **Double**|The y-coordinate of a point on the guide.|

## Return value

Shape


## Remarks

To view guides you add to a master by using the  **AddGuide** method, use the **OpenDrawWindow** method to open the **Master Editing** window.

The following constants declared by the Visio type library are valid values for guide types.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visPoint**|1|Guide point|
| **visHorz**|2|Horizontal guide|
| **visVert**|3|Vertical guide|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]