---
title: Documents.StyleDeleteCanceled Event (Visio)
keywords: vis_sdr.chm10619350
f1_keywords:
- vis_sdr.chm10619350
ms.prod: visio
api_name:
- Visio.Documents.StyleDeleteCanceled
ms.assetid: 9715352b-a34a-1e6f-5080-a922eccc3358
ms.date: 06/08/2017
---


# Documents.StyleDeleteCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelStyleDelete** event.


## Syntax

Private Sub  _expression_ _'StyleDeleteCanceled'(**_ByVal Style As [IVSTYLE]_**)

 _expression_ A variable that represents a [Documents](./Visio.Documents.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[IVSTYLE]**|The style that was going to be deleted.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


