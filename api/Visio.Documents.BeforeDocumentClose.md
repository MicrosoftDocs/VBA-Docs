---
title: Documents.BeforeDocumentClose event (Visio)
keywords: vis_sdr.chm10619025
f1_keywords:
- vis_sdr.chm10619025
ms.prod: visio
api_name:
- Visio.Documents.BeforeDocumentClose
ms.assetid: 62fabfbc-7dcb-990e-ed49-8d8f190bd1eb
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.BeforeDocumentClose event (Visio)

Occurs before a document is closed.


## Syntax

_expression_.**BeforeDocumentClose** (_doc_)

_expression_ A variable that represents a **[Documents](Visio.Documents.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that is going to be closed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]