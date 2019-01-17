---
title: Application.QueryCancelSuspend Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.QueryCancelSuspend
ms.assetid: 1beb9459-f331-d20b-59f0-da505a375a4f
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.QueryCancelSuspend Event (Visio)

Occurs before the operating system enters a suspended state. If any event handler returns  **True** , the Microsoft Visio instance will deny the operating system's request.


## Syntax

Private Sub  _expression_ _'QueryCancelSuspend'(**_ByVal app As [IVAPPLICATION]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that responds to the operating system request.|

## Remarks

 You will typically respond **False** and allow the operating system to enter a suspended state. If you have open network files, you can close them when you receive the **BeforeSuspend** event. If you have open network files that you cannot close, you can return **True** and Visio will deny the operating system's request.




- If any event handler returns  **True** (cancel), the instance fires **SuspendCanceled** and does not enter a suspended state.
    
- If all handlers return  **False** (don't cancel), the instance fires **BeforeSuspend** and then enters a suspended state.
    


If your solution runs outside the Visio process, you cannot be assured of receiving this event. For this reason, you should monitor window messages in your program.

While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


## Example

This VBA macro shows how to capture the  **QueryCancelSuspend** event and allow the operating system to suspend. Declare a **WithEvents** variable to capture events fired by the **Application** object.


```vb
 
Public WithEvents vsoApplication As Visio.Application  
  
Private Function vsoApplication_QueryCancelSuspend(ByVal _ 
    IVisioApplication As IVApplication) As Boolean 
  
    'You agree to let the operating system suspend.  
    vsoApplication_QueryCancelSuspend = False 
  
End Function
```


