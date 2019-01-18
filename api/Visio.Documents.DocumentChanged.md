---
title: Documents.DocumentChanged Event (Visio)
keywords: vis_sdr.chm10619120
f1_keywords:
- vis_sdr.chm10619120
ms.prod: visio
api_name:
- Visio.Documents.DocumentChanged
ms.assetid: 8efdaa32-1c52-fcac-bf5c-fe102774497b
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.DocumentChanged Event (Visio)

Occurs after certain properties of a document are changed.


## Syntax

Private Sub  _expression_ _'DocumentChanged'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents a [Documents](./Visio.Documents.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document whose properties were changed.|

## Remarks

The  **DocumentChanged** event indicates that one of a document's properties, such as **Author** or **Description** , has changed.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


