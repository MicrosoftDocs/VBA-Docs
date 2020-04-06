---
title: Master.ShapeChanged event (Visio)
keywords: vis_sdr.chm10719230
f1_keywords:
- vis_sdr.chm10719230
ms.prod: visio
api_name:
- Visio.Master.ShapeChanged
ms.assetid: e1a2a7bf-bfe1-acfc-ae04-308f9fda7c0a
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.ShapeChanged event (Visio)

Occurs after a property of a shape that is not stored in a cell is changed in a document.


## Syntax

_expression_.**ShapeChanged** (_Shape_)

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape whose property changed.|

## Remarks

To determine which properties have changed when  **ShapeChanged** fires, use the **EventInfo** property. The string returned by the **EventInfo** property contains a list of substrings that identify the properties that changed.

Changes to the following shape properties cause the  **ShapeChanged** event to fire:




-  **Name** (the **EventInfo** property contains "/name")
    
-  **Data1** (the **EventInfo** property contains "/data1")
    
-  **Data2** (the **EventInfo** property contains "/data2")
    
-  **Data3** (the **EventInfo** property contains "/data3")
    
-  **UniqueID** (the **EventInfo** property contains "/uniqueid")
    


If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

If you are handling this event from a program that receives a notification over a connection that was created by using  **AddAdvise**, the _varMoreInfo_ argument to **VisEventProc** contains the string: "/doc=1 /page=1 /shape=Sheet.3"




> [!NOTE] 
> You can use VBA  **WithEvents** variables to sink the **ShapeChanged** event.

For performance considerations, the  **Document** object's event set does not include the **ShapeChanged** event. To sink the **ShapeChanged** event from a **Document** object (and from the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** object in a VBA project), you must use the **AddAdvise** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]