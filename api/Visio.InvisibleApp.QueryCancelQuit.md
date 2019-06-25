---
title: InvisibleApp.QueryCancelQuit event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.QueryCancelQuit
ms.assetid: c0816c40-6118-c64c-7a84-a221debae679
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.QueryCancelQuit event (Visio)

Occurs before the application terminates in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.


## Syntax

_expression_.**QueryCancelQuit** (_app_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio that is going to terminate.|

## Remarks

A Visio instance fires  **QueryCancelQuit** after the user has directed the instance to terminate.




- If any event handler returns  **True** (cancel), the instance fires **QuitCanceled** and does not terminate.
    
- If all handlers return  **False** (don't cancel), the instance fires **BeforeQuit** and then terminates.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]