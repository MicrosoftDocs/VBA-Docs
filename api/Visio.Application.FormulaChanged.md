---
title: Application.FormulaChanged event (Visio)
ms.prod: visio
api_name:
- Visio.Application.FormulaChanged
ms.assetid: f6414b65-cd58-f253-df26-ac33f821799c
ms.date: 06/26/2019
localization_priority: Normal
---


# Application.FormulaChanged event (Visio)

Occurs after a formula changes in a cell in the object that receives the event.


## Syntax

_expression_.**FormulaChanged** (_Cell_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


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
> You can use VBA **WithEvents** variables to sink the **FormulaChanged** event.

For performance considerations, the **Document** object's event set does not include the **FormulaChanged** event. To sink the **FormulaChanged** event from a **Document** object (and the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** object in a VBA project), you must use the **AddAdvise** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]