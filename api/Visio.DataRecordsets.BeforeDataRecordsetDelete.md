---
title: DataRecordsets.BeforeDataRecordsetDelete event (Visio)
keywords: vis_sdr.chm16362030
f1_keywords:
- vis_sdr.chm16362030
ms.prod: visio
api_name:
- Visio.DataRecordsets.BeforeDataRecordsetDelete
ms.assetid: f559b904-df44-3f4e-14d9-6a1493180f27
ms.date: 06/08/2017
localization_priority: Normal
---


# DataRecordsets.BeforeDataRecordsetDelete event (Visio)

Occurs before a **DataRecordset** object is deleted from the **DataRecordsets** collection.


> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

_expression_.**BeforeDataRecordsetDelete** (_DataRecordset_)

_expression_ An expression that returns a **[DataRecordsets](Visio.DataRecordsets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataRecordset_|Required| **[IVDATARECORDSET]**|The data recordset that is going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]