---
title: InvisibleApp.ScreenUpdating Property (Visio)
keywords: vis_sdr.chm17514295
f1_keywords:
- vis_sdr.chm17514295
ms.prod: visio
api_name:
- Visio.InvisibleApp.ScreenUpdating
ms.assetid: c932e8be-b6a6-df5c-d8bc-88aa0c3bd3dc
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.ScreenUpdating Property (Visio)

Determines whether the screen is updated (redrawn) during a series of actions. Read/write.


## Syntax

 _expression_. `ScreenUpdating`

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


## Return value

Integer


## Remarks

Use the  **ScreenUpdating** property to increase performance during a series of actions. For example, you can turn off screen updating while a series of shapes are created so that the screen is not redrawn after each shape appears. Then you can turn screen updating on to update the screen.

If you send a large number of commands to a Microsoft Visio instance while screen updating is turned off, the Visio instance may redisplay the screen occasionally to flush its buffers.

If a program neglects to turn screen updating on after turning it off, the Visio instance turns screen updating back on when a user performs an operation. 


 **Note**  The  **ShowChanges** and **ScreenUpdating** properties are similar in that they are both designed to increase performance during a series of actions, but they work differently. Setting the **ShowChanges** property also sets the **ScreenUpdating** property, but setting the **ScreenUpdating** property does not set the **ShowChanges** property. For a comparison of these two properties, see the **ShowChanges** property.


## Example

This Microsoft Visual Basic code snippet shows how to use the  **ScreenUpdating** property.


```vb
'Turn off screen updating to improve performance during 
'the series of actions that follow. 
 Visio.Application.ScreenUpdating = False 
 
'Drop several shapes. 
'Set the shapes' text. 
'Connect the shapes. 
'Format the connectors. 
 
'Turn screen updating on again when the actions are complete. 
Visio.Application.ScreenUpdating = True 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]