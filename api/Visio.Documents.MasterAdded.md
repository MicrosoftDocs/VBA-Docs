---
title: Documents.MasterAdded Event (Visio)
keywords: vis_sdr.chm10619170
f1_keywords:
- vis_sdr.chm10619170
ms.prod: visio
api_name:
- Visio.Documents.MasterAdded
ms.assetid: aaa83155-bad2-10cb-25cd-a98fc20bc3f0
ms.date: 06/08/2017
---


# Documents.MasterAdded Event (Visio)

Occurs after a new master is added to a document.


## Syntax

Private Sub  _expression_ _'MasterAdded'(**_ByVal Master As [IVMASTER]_**)

 _expression_ A variable that represents a [Documents](./Visio.Documents.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that was added to the document.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


