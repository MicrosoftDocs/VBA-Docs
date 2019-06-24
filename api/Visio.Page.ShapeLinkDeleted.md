---
title: Page.ShapeLinkDeleted event (Visio)
keywords: vis_sdr.chm10962020
f1_keywords:
- vis_sdr.chm10962020
ms.prod: visio
api_name:
- Visio.Page.ShapeLinkDeleted
ms.assetid: e19709c4-45e4-f0f1-8e59-72b1ccbdf130
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.ShapeLinkDeleted event (Visio)

Occurs after the link between a shape and a data row is deleted.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**ShapeLinkDeleted** (_Shape_, _DataRecordsetID_, _DataRowID_)

 _expression_ An expression that returns a **[Page](Visio.Page.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape whose link to a data row was broken.|
| _DataRecordsetID_|Required| **Long**|The ID of the data recordset containing the data row that was linked to the shape.|
| _DataRowID_|Required| **Long**|The ID of the data row that was linked to the shape.|

## Remarks

The  **ShapeLinkDeleted** event is one of a group of events for which the **EventInfo** property of the **Application** object contains extra information.

When the  **ShapeLinkDeleted** event is fired, the **EventInfo** property returns the following string:

 `/DataRecordsetID = n /DataRowID = m`

where  _n_ and _m_ represent the IDs of the data recordset and data row, respectively, associated with the event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]