---
title: Master.PasteToLocation method (Visio)
keywords: vis_sdr.chm10762120
f1_keywords:
- vis_sdr.chm10762120
ms.prod: visio
api_name:
- Visio.Master.PasteToLocation
ms.assetid: c5c94265-23ee-5516-525d-ed3f34d2e7bf
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.PasteToLocation method (Visio)

Pastes a shape to the specified location.


## Syntax

_expression_. `PasteToLocation`( `_xPos_` , `_yPos_` , `_Flags_` )

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _xPos_|Required| **Double**|The x-coordinate at which to place the center of the object's width or PinX, in inches.|
| _yPos_|Required| **Double**|The y-coordinate at which to place the center of the object's height or PinY, in inches.|
| _Flags_|Required| **Long**|The default is zero.|

## Return value

 **Nothing**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]