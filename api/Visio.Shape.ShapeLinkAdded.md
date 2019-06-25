---
title: Shape.ShapeLinkAdded event (Visio)
keywords: vis_sdr.chm11262015
f1_keywords:
- vis_sdr.chm11262015
ms.prod: visio
api_name:
- Visio.Shape.ShapeLinkAdded
ms.assetid: 5cd7431f-18da-184c-7976-06f174cd5f73
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.ShapeLinkAdded event (Visio)

Occurs after a shape is linked to a data row.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**ShapeLinkAdded** (_Shape_, _DataRecordsetID_, _DataRowID_)

 _expression_ An expression that returns a **[Shape](Visio.Shape.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape that is linked to data.|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset containing the data row linked to the shape.|
| _DataRowID_|Required| **Long**|The ID of the data row linked to the shape.|

## Remarks

The  **ShapeLinkAdded** event of the **Shape** object does not occur when a link is created between the source shape itself and a data row. It occurs only when a link is added between one or more sub-shapes of the source shape and a data row or rows.

The  **ShapeLinkAdded** event is one of a group of events for which the **EventInfo** property of the **Application** object contains extra information.

When the  **ShapeLinkAdded** event is fired, the **EventInfo** property returns the following string:

 `/DataRecordsetID = n /DataRowID = m`

where  _n_ and _m_ represent the IDs of the data recordset and data row, respectively, associated with the event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]