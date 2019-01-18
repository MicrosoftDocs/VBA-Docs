---
title: Application.DeferRecalc Property (Visio)
keywords: vis_sdr.chm10013400
f1_keywords:
- vis_sdr.chm10013400
ms.prod: visio
api_name:
- Visio.Application.DeferRecalc
ms.assetid: 25f7ee2e-8987-f03e-5dee-74550bc19f83
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DeferRecalc Property (Visio)

Determines whether the application recalculates cell formulas during a series of actions. Read/write.


## Syntax

 _expression_. `DeferRecalc`

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Return value

Integer


## Remarks

Use the  **DeferRecalc** property to improve performance during a series of actions. For example, you can defer formula recalculation while changing the formulas or values of several cells. When the series of actions is complete, you must always set the **DeferRecalc** property back to the value it had before you changed it. See the following examples.

If you release objects or send a large number of commands to Visio while recalculation is deferred, Visio may at times need to process its queue of pending recalculations. Because of this, use care in setting formulas inside a scope where you want recalculation deferred. Ideally, you should only set formulas when recalculation is turned off.

For example, consider the following Microsoft Visual Basic for Applications (VBA) sequence:




```vb
Dim blsDeferCalcOriginalValue As Boolean 
blsDeferCalcOriginalValue = Application.DeferRecalc 
Application.DeferRecalc = True 
vsoShape.Cells("height").ResultIU = 12 
vsoShape.Cells("width").ResultIU = 14 
Application.DeferRecalc = blsDeferCalcOriginalValue 

```

Because VBA makes and releases a temporary  **Cell** object in the preceding code, Visio will process its queue at that point.

In the following sequence, Visio will not process the recalculation queue until the application turns recalculation on again (or the user performs some operation).




```vb
Dim blsDeferCalcOriginalValue As Boolean 
blsDeferCalcOriginalValue = Application.DeferRecalc 
Application.DeferRecalc = True 
Set vsoCell1 = vsoShape.Cells("Height") 
Set vsoCell2 = vsoShape.Cells("Width") 
vsoCell1.ResultIU = 12 
vsoCell2.ResultIU = 14 
Application.DeferRecalc = blsDeferCalcOriginalValue 

```


