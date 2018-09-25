---
title: InvisibleApp.PageDeleteCanceled Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.PageDeleteCanceled
ms.assetid: 35741231-a4f6-cffb-7004-3c33e538be0b
ms.date: 06/08/2017
---


# InvisibleApp.PageDeleteCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelPageDelete** event.


## Syntax

Private Sub  _expression_ _'PageDeleteCanceled'(**_ByVal Page As [IVPAGE]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Page_|Required| **[IVPAGE]**|The page that was going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


