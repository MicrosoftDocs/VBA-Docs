---
title: Application.BeforeDocumentSave Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.BeforeDocumentSave
ms.assetid: d5d159fb-52e8-2308-6cc2-3b5b4f82fabb
ms.date: 06/08/2017
---


# Application.BeforeDocumentSave Event (Visio)

Occurs before a document is saved.


## Syntax

Private Sub  _expression_ _'BeforeDocumentSave'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that is going to be saved.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


