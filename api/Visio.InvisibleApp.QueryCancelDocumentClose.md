---
title: InvisibleApp.QueryCancelDocumentClose Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.QueryCancelDocumentClose
ms.assetid: 70d38ab1-2468-faa8-85f7-0d2022f314ef
ms.date: 06/08/2017
---


# InvisibleApp.QueryCancelDocumentClose Event (Visio)

Occurs before the application closes a document in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _'QueryCancelDocumentClose'(**_ByVal doc As [IVDOCUMENT]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that is going to be closed.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelDocumentClose** after the user has directed the instance to close a document.




- If any event handler returns  **True** (cancel), the instance fires **DocumentCloseCanceled** and does not close the document.
    
- If all handlers return  **False** (don't cancel), the instance fires **BeforeDocumentClose** and then closes the document.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).


