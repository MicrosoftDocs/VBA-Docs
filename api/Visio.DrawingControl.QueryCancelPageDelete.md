---
title: DrawingControl.QueryCancelPageDelete event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.QueryCancelPageDelete
ms.assetid: c6965236-1e78-736e-484f-b41563e1bdb1
ms.date: 06/08/2017
localization_priority: Normal
---


# DrawingControl.QueryCancelPageDelete event (Visio)

Occurs before the application deletes a page in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.


## Syntax

_expression_.**QueryCancelPageDelete** (_Page_)

_expression_ A variable that represents a **[DrawingControl](Visio.DrawingControl.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Page_|Required| **[IVPAGE]**|The page that is going to be deleted.|

## Remarks

A Visio instance fires  **QueryCancelPageDelete** after the user has directed the instance to delete a page.




- If any event handler returns  **True** (cancel), the instance fires **PageDeleteCanceled** and does not delete the page.
    
- If all handlers return  **False** (don't cancel) the instance fires **BeforePageDelete** and then deletes the page.
    


While a Microsoft Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]