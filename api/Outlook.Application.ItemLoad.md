---
title: Application.ItemLoad event (Outlook)
keywords: vbaol11.chm446
f1_keywords:
- vbaol11.chm446
ms.prod: outlook
api_name:
- Outlook.Application.ItemLoad
ms.assetid: aed0656d-4e5a-550a-1116-76773215a897
ms.date: 09/19/2018
localization_priority: Normal
---


# Application.ItemLoad event (Outlook)

Occurs when an Outlook item is loaded into memory.


## Syntax

_expression_. ItemLoad( _Item_ )

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|A weak object reference for the loaded Outlook item.|

## Remarks

This event occurs when the Outlook item begins to load into memory. Data for the item is not yet available, other than the values for the **Class** and **MessageClass** properties of the Outlook item, so an error occurs when calling any property other than **Class** or **MessageClass** for the Outlook item returned in _Item_. 

Similarly, an error occurs if you attempt to call any method from the Outlook item, or if you call the **[GetObjectReference](Outlook.Application.GetObjectReference.md)** method of the **[Application](Outlook.Application.md)** object on the Outlook item returned in _Item_.

The **ItemLoad** event should typically be implemented as a means to hook up item-level event handlers such as **BeforeRead**, **Open**, **Send**, and **Write**.

> [!WARNING] 
> The _Item_ object passed in this event should not be cached for any use outside the scope of this event.

This event is not raised when the following conditions occur:

- An Outlook item is synchronized with a folder.
    
- A server-side rule is triggered for an Outlook item.
    
- A reminder is triggered for an Outlook item.
    
- A Desktop Alert is displayed for an Outlook item.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]