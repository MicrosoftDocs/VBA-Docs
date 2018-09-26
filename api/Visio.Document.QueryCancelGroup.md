---
title: Document.QueryCancelGroup Event (Visio)
keywords: vis_sdr.chm10562000
f1_keywords:
- vis_sdr.chm10562000
ms.prod: visio
api_name:
- Visio.Document.QueryCancelGroup
ms.assetid: 0fb4f654-f501-32d7-d94d-5240cfc82eb4
ms.date: 06/08/2017
---


# Document.QueryCancelGroup Event (Visio)

Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _'QueryCancelGroup'(**_ByVal Selection As [IVSELECTION]_**)

 _expression_ A variable that represents a [Document](./Visio.Document.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that is going to be grouped.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelGroup** after the user has directed the instance to group a selection of shapes.




- If any event handler returns  **True** (cancel), the instance fires **GroupCanceled** and does not group the shapes.
    
- If all handlers return  **False** (do not cancel), the grouping is performed.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you use Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


