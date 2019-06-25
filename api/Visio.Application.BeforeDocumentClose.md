---
title: Application.BeforeDocumentClose event (Visio)
ms.prod: visio
api_name:
- Visio.Application.BeforeDocumentClose
ms.assetid: c0d7815e-25bb-7b7e-f80b-81472edc47ca
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.BeforeDocumentClose event (Visio)

Occurs before a document is closed.


## Syntax

_expression_.**BeforeDocumentClose** (_doc_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


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

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this event maps to the following types:

- **Microsoft.Office.Interop.Visio.EApplication_BeforeDocumentCloseEventHandler** (the **BeforeDocumentClose** delegate)   
- **Microsoft.Office.Interop.Visio.EApplication_Event.BeforeDocumentClose** (the **BeforeDocumentClose** event)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]