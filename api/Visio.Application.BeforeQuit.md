---
title: Application.BeforeQuit event (Visio)
ms.prod: visio
api_name:
- Visio.Application.BeforeQuit
ms.assetid: ad5ed704-4e7e-f8a9-b238-3c552dc3f292
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.BeforeQuit event (Visio)

Occurs before a Microsoft Visio instance terminates.


## Syntax

_expression_.**BeforeQuit** (_app_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that is going to terminate.|

## Remarks

When programming with Microsoft Visual Basic, use the **BeforeDocumentClose** event instead of the **BeforeQuit** event. The code in a Visual Basic project of a Visio document never has the chance to respond to the **BeforeQuit** event because the project is a property of a document, and all documents are closed before the **BeforeQuit** event notification is sent.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]