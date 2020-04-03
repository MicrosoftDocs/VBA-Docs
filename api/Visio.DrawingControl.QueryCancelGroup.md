---
title: DrawingControl.QueryCancelGroup event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.QueryCancelGroup
ms.assetid: 630abedc-0b1a-8ad4-47d7-51215c1f0c43
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.QueryCancelGroup event (Visio)

Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.


## Syntax

_expression_.**QueryCancelGroup** (_Selection_)

_expression_ A variable that represents a **[DrawingControl](Visio.DrawingControl.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that is going to be grouped.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelGroup** after the user has directed the instance to group a selection of shapes.




- If any event handler returns  **True** (cancel), the instance fires **GroupCanceled** and does not group the shapes.
    
- If all handlers return  **False** (don't cancel), the grouping is performed.
    


In some cases, such as when a shape that has a  **ForeignType** property of **visTypeMetafile** is grouped, the initial shape will be deleted and replaced with new shapes. In such cases, the Visio instance will subsequently fire **BeforeSelectionDelete** and **BeforeShapeDelete** events before grouping the shapes.

While a Visio instance is firing a query or cancel event, it will respond to inquiries from client code but will refuse to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]