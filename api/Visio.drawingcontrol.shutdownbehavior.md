---
title: DrawingControl.ShutDownBehavior property (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.ShutDownBehavior
ms.assetid: 19c3e160-4b1d-40f1-b41d-69f21fca1d0d
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.ShutDownBehavior property (Visio)

Determines how the Visio Drawing Control unloads the Visio application when the  **DrawingControl** object is released. Read/write **Integer**.


## Syntax

_expression_.**ShutDownBehavior**

_expression_ A variable that represents a **[DrawingControl](Visio.DrawingControl.md)** object.


## Return value

 **Integer**


## Remarks

A value of 0 (the default) does not unload MSO dlls when the drawing control is released. A value of 1 unloads the Visio application and MSO dlls when the drawing control is released.


## See also


[DrawingControl Object](Visio.DrawingControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]