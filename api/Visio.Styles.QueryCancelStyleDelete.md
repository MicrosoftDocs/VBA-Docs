---
title: Styles.QueryCancelStyleDelete Event (Visio)
keywords: vis_sdr.chm11519305
f1_keywords:
- vis_sdr.chm11519305
ms.prod: visio
api_name:
- Visio.Styles.QueryCancelStyleDelete
ms.assetid: 08d1f26c-bdcd-1f5d-3dd6-a657ef6619ab
ms.date: 06/08/2017
---


# Styles.QueryCancelStyleDelete Event (Visio)

Occurs before the application deletes a style in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _'QueryCancelStyleDelete'(**_ByVal Style As [IVSTYLE]_**)

 _expression_ A variable that represents a [Styles](./Visio.Styles.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _style_|Required| **[IVSTYLE]**|The style that is going to be deleted.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelStyleDelete** after the user has directed the instance to delete a style.




- If any event handler returns  **True** (cancel), the instance fires **StyleDeleteCanceled** and does not delete the style.
    
- If all handlers return  **False** (don't cancel), the instance fires **BeforeStyleDelete** and then deletes the style.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


