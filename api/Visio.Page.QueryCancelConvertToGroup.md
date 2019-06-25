---
title: Page.QueryCancelConvertToGroup event (Visio)
keywords: vis_sdr.chm10919325
f1_keywords:
- vis_sdr.chm10919325
ms.prod: visio
api_name:
- Visio.Page.QueryCancelConvertToGroup
ms.assetid: a9dc79ef-2a4c-398a-4bf3-d29e0cf916f4
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.QueryCancelConvertToGroup event (Visio)

Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.


## Syntax

_expression_.**QueryCancelConvertToGroup** (_Selection_)

_expression_ A variable that represents a **[Page](Visio.Page.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that is going to be converted to a group.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelConvertToGroup** after the user has directed the instance to convert one or more shapes into groups.




- If any event handler returns  **True** (cancel), the instance fires **ConvertToGroupCanceled** and does not convert the shapes.
    
- If all handlers return  **False** (don't cancel), the conversion is performed.
    


In some cases, such as when a shape that has a  **ForeignType** property of **visTypeMetafile** is converted to a group, the initial shape is deleted and replaced with new shapes. In such cases, the Visio instance subsequently fires **BeforeSelectionDelete** and **BeforeShapeDelete** events before converting the shapes.

While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]