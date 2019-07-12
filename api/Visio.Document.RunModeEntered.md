---
title: Document.RunModeEntered event (Visio)
keywords: vis_sdr.chm10519210
f1_keywords:
- vis_sdr.chm10519210
ms.prod: visio
api_name:
- Visio.Document.RunModeEntered
ms.assetid: 8e582dd1-b2c5-72e5-b144-510726d35a18
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.RunModeEntered event (Visio)

Occurs after a document enters run mode.


## Syntax

_expression_.**RunModeEntered** (_doc_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that entered run mode.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]