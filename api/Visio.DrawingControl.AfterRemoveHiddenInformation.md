---
title: DrawingControl.AfterRemoveHiddenInformation event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.AfterRemoveHiddenInformation
ms.assetid: 53c41981-3f24-53e3-dea5-204e0ad6f046
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.AfterRemoveHiddenInformation event (Visio)

Occurs when hidden information is removed from the document.


## Syntax

_expression_.**AfterRemoveHiddenInformation** (_doc_)

_expression_ An expression that returns a **[DrawingControl](Visio.DrawingControl.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document from which hidden information has been removed.|

## Remarks

The **AfterRemoveHiddenInformation** event is one of a group of events for which the **EventInfo** property of the **Application** object contains extra information.

When the **AfterRemoveHiddenInformation** event is fired, the **EventInfo** property returns a string that contains information about which items were removed from the document, consisting of the sum of applicable constant values from the **[VisRemoveHiddenInfoItems](Visio.visremovehiddeninfoitems.md)** enumeration.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]