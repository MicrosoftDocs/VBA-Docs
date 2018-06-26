---
title: Application.DataRecordsetAdded Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.DataRecordsetAdded
ms.assetid: 04a54ec4-6f87-ac4d-f35c-bc3debca4a65
ms.date: 06/08/2017
---


# Application.DataRecordsetAdded Event (Visio)

Occurs when a  **DataRecordset** object is added to a **DataRecordsets** collection.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

Private Sub  _expression_ _'DataRecordsetAdded'( **_ByVal DataRecordset As [IVDATARECORDSET]_** )

 _expression_ An expression that returns a [Application](./Visio.Application.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataRecordset_|Required| **[IVDATARECORDSET]**|The data recordset that was added.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](../visio/Concepts/event-codesvisio.md).


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


