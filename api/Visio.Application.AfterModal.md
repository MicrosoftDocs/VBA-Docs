---
title: Application.AfterModal Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.AfterModal
ms.assetid: e19a0ef3-349c-1d7f-9856-7ef6c66f5f0e
ms.date: 06/08/2017
---


# Application.AfterModal Event (Visio)

Occurs after the Microsoft Visio instance leaves a modal state.


## Syntax

Private Sub  _expression_ _'AfterModal'(**_ByVal app As [IVAPPLICATION]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance that is no longer modal.|

## Remarks

Visio becomes modal when it displays a dialog box. A modal instance of Visio does not handle Automation calls. The  **BeforeModal** event indicates that the instance is about to become modal, and the **AfterModal** event indicates that the instance is no longer modal.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


