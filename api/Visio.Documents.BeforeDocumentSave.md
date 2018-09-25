---
title: Documents.BeforeDocumentSave Event (Visio)
keywords: vis_sdr.chm10619030
f1_keywords:
- vis_sdr.chm10619030
ms.prod: visio
api_name:
- Visio.Documents.BeforeDocumentSave
ms.assetid: 7d678fb6-20eb-b976-19dc-f97f32e7f466
ms.date: 06/08/2017
---


# Documents.BeforeDocumentSave Event (Visio)

Occurs before a document is saved.


## Syntax

Private Sub  _expression_ _'BeforeDocumentSave'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents a [Documents](./Visio.Documents.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that is going to be saved.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


