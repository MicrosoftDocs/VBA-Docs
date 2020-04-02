---
title: DataRecordsetChangedEvent object (Visio)
keywords: vis_sdr.chm61025
f1_keywords:
- vis_sdr.chm61025
ms.prod: visio
api_name:
- Visio.DataRecordsetChangedEvent
ms.assetid: 3575c6f6-081d-4632-d720-efad1c977a9a
ms.date: 06/19/2019
localization_priority: Normal
---


# DataRecordsetChangedEvent object (Visio)

Passed by Microsoft Visio as the _pSubjectObj_ argument to the **[VisEventProc](Visio.IVisEventProc.VisEventProc.md)** method of the **IVisEventProc** interface when events related to refreshing a data recordset fire.

> [!NOTE] 
> This Visio object or member is available only to licensed users of Visio Professional 2013.


## Remarks

To handle event notifications in Visio, create a class module that implements the **IVisEventProc** interface in Microsoft Visual Basic for Applications (VBA) or Microsoft Visual Basic, and then create an instance of this class to pass as an argument to the **[AddAdvise](Visio.EventList.AddAdvise.md)** method of the **EventList** collection.

When data recordset rows are added, changed, or deleted, and when data recordset columns are added or deleted, in each case as a result of a data recordset being refreshed, properties of the **DataRecordsetChangedEvent** object that is passed to the **VisEventProc** method return arrays of the affected rows or columns.

All properties of the **DataRecordsetChangedEvent** object are read-only.

## Properties

- [Application](Visio.DataRecordsetChangedEvent.Application.md)
- [DataColumnsAdded](Visio.DataRecordsetChangedEvent.DataColumnsAdded.md)
- [DataColumnsChanged](Visio.DataRecordsetChangedEvent.DataColumnsChanged.md)
- [DataColumnsDeleted](Visio.DataRecordsetChangedEvent.DataColumnsDeleted.md)
- [DataRecordset](Visio.DataRecordsetChangedEvent.DataRecordset.md)
- [DataRowsAdded](Visio.DataRecordsetChangedEvent.DataRowsAdded.md)
- [DataRowsDeleted](Visio.DataRecordsetChangedEvent.DataRowsDeleted.md)
- [ObjectType](Visio.DataRecordsetChangedEvent.ObjectType.md)
- [Stat](Visio.DataRecordsetChangedEvent.Stat.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]