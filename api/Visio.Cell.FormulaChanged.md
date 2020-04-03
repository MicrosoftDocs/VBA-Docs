---
title: Cell.FormulaChanged event (Visio)
keywords: vis_sdr.chm10119160
f1_keywords:
- vis_sdr.chm10119160
ms.prod: visio
api_name:
- Visio.Cell.FormulaChanged
ms.assetid: 7f612470-ea40-1b7e-7334-825a124a96f3
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.FormulaChanged event (Visio)

Occurs after a formula changes in a cell in the object that receives the event.


## Syntax

_expression_.**FormulaChanged** (_Cell_)

_expression_ A variable that represents a **[Cell](Visio.Cell.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cell_|Required| **[IVCELL]**|The cell whose formula changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).




> [!NOTE] 
> You can use VBA  **WithEvents** variables to sink the **FormulaChanged** event.

For performance considerations, the  **Document** object's event set does not include the **FormulaChanged** event. To sink the **FormulaChanged** event from a **Document** object (and the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** object in a VBA project), you must use the **AddAdvise** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]