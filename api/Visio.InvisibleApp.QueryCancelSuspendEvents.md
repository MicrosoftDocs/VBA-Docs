---
title: InvisibleApp.QueryCancelSuspendEvents Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.QueryCancelSuspendEvents
ms.assetid: 375763d4-fbb8-fa08-8fcd-bf5dc80aceb9
ms.date: 06/08/2017
localization_priority: Normal
---


# InvisibleApp.QueryCancelSuspendEvents Event (Visio)

Occurs before the application suspends events in response to a client code. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _'QueryCancelSuspendEvents'(**_ByVal app As [IVAPPLICATION]_**)

 _expression_ An expression that returns a [InvisibleApp](./Visio.InvisibleApp.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio in which firing of events is going to be suspended.|

## Return value

nothing


## Remarks

A Visio instance fires  **QueryCancelSuspendEvents** after client code has directed the instance to suspend events.




- If any event handler returns  **True** (cancel), the instance fires **SuspendEventsCanceled** and does not suspend events.
    
- If all handlers return  **False** (don't cancel), the suspension occurs.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]