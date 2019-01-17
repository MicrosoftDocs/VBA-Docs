---
title: Pages.QueryCancelSelectionDelete Event (Visio)
keywords: vis_sdr.chm11019320
f1_keywords:
- vis_sdr.chm11019320
ms.prod: visio
api_name:
- Visio.Pages.QueryCancelSelectionDelete
ms.assetid: d9749c36-d336-f251-7f69-c48bcf590d56
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.QueryCancelSelectionDelete Event (Visio)

Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _'QueryCancelSelectionDelete'(**_ByVal Selection As [IVSELECTION]_**)

 _expression_ A variable that represents a [Pages](./Visio.Pages.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that is going to be deleted.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelSelectionDelete** after the user has directed the instance to delete one or more shapes.




- If any event handler returns  **True** (cancel), the instance fires **SelectionDeleteCanceled** and does not delete the shapes.
    
- If all handlers return **False** (don't cancel), the instance fires **BeforeSelectionDelete** and **BeforeShapeDelete** and then deletes the shapes.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]