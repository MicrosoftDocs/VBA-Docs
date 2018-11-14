---
title: Application.DataRecordsetChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.DataRecordsetChanged
ms.assetid: 8be61b1a-3a3c-5880-47bc-e2cac9bb60f9
ms.date: 06/08/2017
---


# Application.DataRecordsetChanged Event (Visio)

Occurs when a data recordset changes as a result of being refreshed.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

Private Sub  _expression_ _'DataRecordsetChanged'(**_ByVal DataRecordsetChanged As IVDATARECORDSETCHANGEDEVENT_**)

 _expression_ An expression that returns a [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataRecordsetChanged_|Required| **[IVDATARECORDSETCHANGEDEVENT]**|The data recordset that changed.|

## Remarks

When the  **DataRecordsetChanged** event fires, the **[DataRecordsetChangedEvent](Visio.DataRecordsetChangedEvent.md)** object is passed to the **[IVisEventProc.VisEventProc](Visio.IVisEventProc.VisEventProc.md)** method as the pSubjectObj parameter, which represents he subject of the event?the object to which the event occurs.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


