---
title: Masters.QueryCancelMasterDelete event (Visio)
keywords: vis_sdr.chm10819310
f1_keywords:
- vis_sdr.chm10819310
ms.prod: visio
api_name:
- Visio.Masters.QueryCancelMasterDelete
ms.assetid: 69aa351f-2e89-545d-0cf8-f650d532d3a6
ms.date: 06/08/2017
localization_priority: Normal
---


# Masters.QueryCancelMasterDelete event (Visio)

Occurs before the application deletes a master in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.


## Syntax

_expression_.**QueryCancelMasterDelete** (_Master_)

_expression_ A variable that represents a **[Masters](Visio.Masters.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that is going to be deleted.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelMasterDelete** after the user has directed the instance to delete a master.




- If any event handler returns  **True** (cancel), the instance fires **MasterDeleteCanceled** and does not delete the master.
    
- If all handlers return  **False** (don't cancel), the instance fires **BeforeMasterDelete** and then deletes the master.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]