---
title: DrawingControl.DocumentCloseCanceled Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.DocumentCloseCanceled
ms.assetid: 0de2255a-49c6-88cf-04f0-7e32705c329e
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.DocumentCloseCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelDocumentClose** event.


## Syntax

Private Sub  _expression_ _'DocumentCloseCanceled'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents a [DrawingControl](./Visio.DrawingControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that was going to be closed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


