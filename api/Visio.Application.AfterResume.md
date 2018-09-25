---
title: Application.AfterResume Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.AfterResume
ms.assetid: 73cac713-6559-ae7c-32a6-5c421302a3d9
ms.date: 06/08/2017
---


# Application.AfterResume Event (Visio)

Occurs when the operating system resumes normal operation after having been suspended.


## Syntax

 Private Sub _expression_ _'AfterResume'(**_ByVal app As [IVAPPLICATION]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that resumes after the operating system resumes normal operation.|

## Remarks

You can use the  **AfterResume** event to reopen any network files that you may have closed in response to the **BeforeSuspend** event.

If your solution runs outside the Microsoft Visio process, you cannot be assured of receiving this event. For this reason, you should monitor window messages in your program.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


