---
title: Document.DesignModeEntered Event (Visio)
keywords: vis_sdr.chm10519110
f1_keywords:
- vis_sdr.chm10519110
ms.prod: visio
api_name:
- Visio.Document.DesignModeEntered
ms.assetid: c8fc31b5-8770-f068-d469-aeb110214824
ms.date: 06/08/2017
---


# Document.DesignModeEntered Event (Visio)

Occurs before a document enters design mode.


## Syntax

Private Sub  _expression_ _'DesignModeEntered'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents a [Document](./Visio.Document.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that is going to enter design mode.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


