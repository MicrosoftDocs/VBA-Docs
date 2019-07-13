---
title: Document.QueryCancelUngroup event (Visio)
keywords: vis_sdr.chm10519330
f1_keywords:
- vis_sdr.chm10519330
ms.prod: visio
api_name:
- Visio.Document.QueryCancelUngroup
ms.assetid: e25505a9-a2ae-dc68-8bf6-ac4252c7f5e6
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.QueryCancelUngroup event (Visio)

Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns **True**, the operation is canceled.


## Syntax

_expression_.**QueryCancelUngroup** (_Selection_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that is going to be ungrouped.|

## Remarks

A Microsoft Visio instance fires **QueryCancelUngroup** after the user has directed the instance to ungroup one or more shapes.




- If any event handler returns **True** (cancel), the instance fires **UngroupCanceled** and does not ungroup the shapes.
    
- If all handlers return **False** (don't cancel), the instance fires **ShapeParentChanged**, **BeforeSelectionDelete**, and **BeforeShapeDelete**, and then ungroups the shapes.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]