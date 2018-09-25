---
title: InvisibleApp.BeforeDocumentClose Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.BeforeDocumentClose
ms.assetid: 663dab28-2ea5-30ad-a694-5c2bc75210fb
ms.date: 06/08/2017
---


# InvisibleApp.BeforeDocumentClose Event (Visio)

Occurs before a document is closed.


## Syntax

Private Sub  _expression_ _'BeforeDocumentClose'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that is going to be closed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


