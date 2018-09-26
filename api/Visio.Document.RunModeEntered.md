---
title: Document.RunModeEntered Event (Visio)
keywords: vis_sdr.chm10519210
f1_keywords:
- vis_sdr.chm10519210
ms.prod: visio
api_name:
- Visio.Document.RunModeEntered
ms.assetid: 8e582dd1-b2c5-72e5-b144-510726d35a18
ms.date: 06/08/2017
---


# Document.RunModeEntered Event (Visio)

Occurs after a document enters run mode.


## Syntax

Private Sub  _expression_ _'RunModeEntered'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents a [Document](./Visio.Document.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that entered run mode.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


