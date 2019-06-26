---
title: Application.SelectionAdded event (Visio)
ms.prod: visio
api_name:
- Visio.Application.SelectionAdded
ms.assetid: 76a6c2c6-c2ab-97ad-a906-0780a81f7eb2
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SelectionAdded event (Visio)

Occurs after one or more shapes are added to a document.


## Syntax

_expression_.**SelectionAdded** (_Selection_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that was added to the document.|

## Remarks

A **Shape** object can serve as the source object for the **SelectionAdded** event if the shape's **Type** property is **visTypeGroup** (2) or **visTypePage** (1).

The **SelectionAdded** and **ShapeAdded** events are similar in that they both fire after shape(s) are created. They differ in how they behave when a single operation adds several shapes. Suppose a **Paste** operation creates three new shapes. The **ShapeAdded** event fires three times and acts on each of the three objects. The **SelectionAdded** event fires once, and it acts on a **Selection** object in which the three new shapes are selected.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).




> [!NOTE] 
> You can use VBA **WithEvents** variables to sink the **SelectionAdded** event.

For performance considerations, the **Document** object's event set does not include the **SelectionAdded** event. To sink the **SelectionAdded** event from a **Document** object (and the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** object in a VBA project), you must use the **AddAdvise** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]