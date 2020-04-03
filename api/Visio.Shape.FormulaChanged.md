---
title: Shape.FormulaChanged event (Visio)
keywords: vis_sdr.chm11219160
f1_keywords:
- vis_sdr.chm11219160
ms.prod: visio
api_name:
- Visio.Shape.FormulaChanged
ms.assetid: cf141b03-5eaf-bf42-a64f-049f8dec2a14
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.FormulaChanged event (Visio)

Occurs after a formula changes in a cell in the object that receives the event.


## Syntax

_expression_.**FormulaChanged** (_Cell_)

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


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