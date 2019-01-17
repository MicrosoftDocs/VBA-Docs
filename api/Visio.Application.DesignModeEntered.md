---
title: Application.DesignModeEntered Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.DesignModeEntered
ms.assetid: 312f0bda-1375-e176-f5c5-4ebd3c9c8b6d
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DesignModeEntered Event (Visio)

Occurs before a document enters design mode.


## Syntax

Private Sub  _expression_ _'DesignModeEntered'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that is going to enter design mode.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]