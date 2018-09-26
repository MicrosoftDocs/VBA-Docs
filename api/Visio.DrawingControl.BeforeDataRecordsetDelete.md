---
title: DrawingControl.BeforeDataRecordsetDelete Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.BeforeDataRecordsetDelete
ms.assetid: 70e30b15-6254-b12b-6f46-ce1f7ae07140
ms.date: 06/08/2017
---


# DrawingControl.BeforeDataRecordsetDelete Event (Visio)

Occurs before a  **DataRecordset** object is deleted from the **DataRecordsets** collection.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

Private Sub  _expression_ _'BeforeDataRecordsetDelete'(**_ByVal DataRecordset As IVDATARECORDSET_**)

 _expression_ An expression that returns a [DrawingControl](./Visio.DrawingControl.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataRecordset_|Required| **[IVDATARECORDSET]**|The data recordset that is going to be deleted.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


