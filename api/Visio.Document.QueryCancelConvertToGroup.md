---
title: Document.QueryCancelConvertToGroup Event (Visio)
keywords: vis_sdr.chm10519325
f1_keywords:
- vis_sdr.chm10519325
ms.prod: visio
api_name:
- Visio.Document.QueryCancelConvertToGroup
ms.assetid: 1199e5ac-26b5-c5ca-106f-1ff4b833b933
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.QueryCancelConvertToGroup Event (Visio)

Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _'QueryCancelConvertToGroup'(**_ByVal Selection As [IVSELECTION]_**)

 _expression_ A variable that represents a [Document](./Visio.Document.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that is going to be converted to a group.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelConvertToGroup** after the user has directed the instance to convert one or more shapes into groups.




- If any event handler returns  **True** (cancel), the instance fires **ConvertToGroupCanceled** and does not convert the shapes.
    
- If all handlers return  **False** (don't cancel), the conversion is performed.
    


In some cases, such as when a shape that has a  **ForeignType** property of **visTypeMetafile** is converted to a group, the initial shape is deleted and replaced with new shapes. In such cases, the Visio instance subsequently fires **BeforeSelectionDelete** and **BeforeShapeDelete** events before converting the shapes.

While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


