---
title: Application.QueryCancelStyleDelete event (Visio)
ms.prod: visio
api_name:
- Visio.Application.QueryCancelStyleDelete
ms.assetid: 7f3ce781-67d8-7a6e-d8f0-b077c8956b12
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.QueryCancelStyleDelete event (Visio)

Occurs before the application deletes a style in response to a user action in the interface. If any event handler returns **True**, the operation is canceled.


## Syntax

_expression_.**QueryCancelStyleDelete** (_Style_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[IVSTYLE]**|The style that is going to be deleted.|

## Remarks

A Microsoft Visio instance fires **QueryCancelStyleDelete** after the user has directed the instance to delete a style.

- If any event handler returns **True** (cancel), the instance fires **StyleDeleteCanceled** and does not delete the style.
    
- If all handlers return **False** (don't cancel), the instance fires **BeforeStyleDelete** and then deletes the style.
    
While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]