---
title: Shape.SelectionDeleteCanceled Event (Visio)
keywords: vis_sdr.chm11219365
f1_keywords:
- vis_sdr.chm11219365
ms.prod: visio
api_name:
- Visio.Shape.SelectionDeleteCanceled
ms.assetid: 10811705-9619-d4d8-80f5-f1fa08eed52f
ms.date: 06/08/2017
---


# Shape.SelectionDeleteCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.


## Syntax

Private Sub  _expression_ _'SelectionDeleteCanceled'(**_ByVal Selection As [IVSELECTION]_**)

 _expression_ A variable that represents a [Shape](./Visio.Shape.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that was going to be deleted.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


