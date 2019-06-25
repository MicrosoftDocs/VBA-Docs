---
title: InvisibleApp.BeforePageDelete event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.BeforePageDelete
ms.assetid: dc7a6fde-1794-b91f-0990-391a78c6039c
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.BeforePageDelete event (Visio)

Occurs before a page is deleted.


## Syntax

_expression_.**BeforePageDelete** (_Page_)

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Page_|Required| **[IVPAGE]**|The page that is going to be deleted.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]