---
title: Documents.UngroupCanceled Event (Visio)
keywords: vis_sdr.chm10619375
f1_keywords:
- vis_sdr.chm10619375
ms.prod: visio
api_name:
- Visio.Documents.UngroupCanceled
ms.assetid: ac75485f-bd09-8339-d403-8587230a99c5
ms.date: 06/08/2017
---


# Documents.UngroupCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.


## Syntax

Private Sub  _expression_ _'UngroupCanceled'(**_ByVal Selection As [IVSELECTION]_**)

 _expression_ A variable that represents a [Documents](./Visio.Documents.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that was going to be ungrouped.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


