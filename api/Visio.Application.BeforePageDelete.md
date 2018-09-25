---
title: Application.BeforePageDelete Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.BeforePageDelete
ms.assetid: 658e3367-2f5b-e2d4-6c07-9b4463ee500a
ms.date: 06/08/2017
---


# Application.BeforePageDelete Event (Visio)

Occurs before a page is deleted.


## Syntax

Private Sub  _expression_ _'BeforePageDelete'(**_ByVal Page As [IVPAGE]_**)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Page_|Required| **[IVPAGE]**|The page that is going to be deleted.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


