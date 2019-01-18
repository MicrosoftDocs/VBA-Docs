---
title: WebBrowserControl.Updated event (Access)
keywords: vbaac10.chm143123
f1_keywords:
- vbaac10.chm143123
ms.prod: access
api_name:
- Access.WebBrowserControl.Updated
ms.assetid: 844fb989-98a5-6588-77e3-789bb9a7112c
ms.date: 06/08/2017
localization_priority: Normal
---


# WebBrowserControl.Updated event (Access)

The  **Updated** event occurs when an OLE object's data has been modified.


## Syntax

_expression_. `Updated`( ` _Code_`, )

_expression_ A variable that represents a [WebBrowserControl](Access.WebBrowserControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Code_|Required|**Integer**||

## Return value

nothing


## Remarks

To run a macro or event procedure when this event occurs, set the  **OnUpdated** property to the name of the macro or to [Event Procedure].

You can use this event to determine if an object's data has been changed since it was last saved.

The  **Updated** event occurs when the data in an OLE object has been modified. This update can come from the application in which the object was created or from one of the linked copies of this object. As a result, this event is asynchronous with other Microsoft Access control events.


 **Note**  The  **Updated** event and the **BeforeUpdate** and **AfterUpdate** events for bound and unbound object frames are not related. The **Updated** event occurs when an OLE object's data is changed, and the **BeforeUpdate** and **AfterUpdate** events occur when data is updated. Although not related, all three events usually occur when an OLE object's data is changed. The **Updated** event generally occurs before the **BeforeUpdate** and **AfterUpdate** events; however, this may not happen every time.


## See also


[WebBrowserControl Object](Access.WebBrowserControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]