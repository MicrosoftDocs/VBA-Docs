---
title: DrawingControl.BeforeDocumentClose event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.BeforeDocumentClose
ms.assetid: f2cf89e1-f9cc-cc59-0a89-c7214a5975ef
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.BeforeDocumentClose event (Visio)

Occurs before a document is closed.


## Syntax

_expression_.**BeforeDocumentClose** (_doc_)

_expression_ A variable that represents a **[DrawingControl](Visio.DrawingControl.md)** object.


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