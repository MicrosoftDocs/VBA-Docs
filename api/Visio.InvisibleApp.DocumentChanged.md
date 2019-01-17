---
title: InvisibleApp.DocumentChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.DocumentChanged
ms.assetid: d822ab40-99a5-d308-d820-a8834f65fee8
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.DocumentChanged Event (Visio)

Occurs after certain properties of a document are changed.


## Syntax

Private Sub  _expression_ _'DocumentChanged'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document whose properties were changed.|

## Remarks

The  **DocumentChanged** event indicates that one of a document's properties, such as **Author** or **Description** , has changed.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]