---
title: DataRecordset.EventList property (Visio)
keywords: vis_sdr.chm16460610
f1_keywords:
- vis_sdr.chm16460610
ms.prod: visio
api_name:
- Visio.DataRecordset.EventList
ms.assetid: 419cdd3d-cb12-cbb6-5e47-d343b1a84d74
ms.date: 06/08/2017
localization_priority: Normal
---


# DataRecordset.EventList property (Visio)

Returns the **[EventList](Visio.EventList.md)** collection of the **DataRecordset** object. Read-only.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**EventList**

_expression_ An expression that returns a **[DataRecordset](Visio.DataRecordset.md)** object.


## Return value

EventList


## Remarks

Once you retrieve the **EventList** collection, to receive a notification when one of the events in that collection fires, you can pass the ID of the **[Event](Visio.Event.md)** object that represents that event to the **[EventList.AddAdvise](Visio.EventList.AddAdvise.md)** method as its EventCode parameter.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]