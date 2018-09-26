---
title: Application.ShapeChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.ShapeChanged
ms.assetid: aac5dfc5-370e-8299-4e3e-39fe9a7000d2
ms.date: 06/08/2017
---


# Application.ShapeChanged Event (Visio)

Occurs after a property of a shape that is not stored in a cell is changed in a document.


## Syntax

Private Sub  _expression_ _'ShapeChanged'(**_ByVal Shape As [IVSHAPE]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


### Parameters



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
    


If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

If you are handling this event from a program that receives a notification over a connection that was created by using  **AddAdvise** , the _varMoreInfo_ argument to **VisEventProc** contains the string: "/doc=1 /page=1 /shape=Sheet.3"




 **Note**  You can use VBA  **WithEvents** variables to sink the **ShapeChanged** event.

For performance considerations, the  **Document** object's event set does not include the **ShapeChanged** event. To sink the **ShapeChanged** event from a **Document** object (and from the **ThisDocument** object in a VBA project), you must use the **AddAdvise** method.


