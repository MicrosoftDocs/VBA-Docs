---
title: Application.AdvancedSearchStopped event (Outlook)
keywords: vbaol11.chm436
f1_keywords:
- vbaol11.chm436
ms.prod: outlook
api_name:
- Outlook.Application.AdvancedSearchStopped
ms.assetid: a1a4ec9f-c0e3-6acd-b63c-89194ed70efd
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AdvancedSearchStopped event (Outlook)

Occurs when a specified  **[Search](Outlook.Search.md)** object's **[Stop](Outlook.Search.Stop.md)** method has been executed.


## Syntax

_expression_. `AdvancedSearchStopped`( `_SearchObject_` )

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SearchObject_|Required| **Search**|The  **[Search](Outlook.Search.md)** object returned by the **[AdvancedSearch](Outlook.Application.AdvancedSearch.md)** method.|

## Remarks

After this event is fired, the  **Search** object's **[Results](Outlook.Results.md)** collection will no longer be updated. This event can only be triggered programmatically.


## Example

The following Visual Basic for Applications (VBA) example starts searching the  **Inbox** for items with subject equal to "Test" and immediately stops the search. This causes the `AdvanceSearchStopped` event procedure to be run. The sample code must be placed in a class module such as `ThisOutlookSession`. The  `StopSearch()` procedure must be called before the event procedure can be called by Microsoft Outlook.


```vb
Sub StopSearch() 
 
 Dim sch As Outlook.Search 
 
 Dim strScope As String 
 
 Dim strFilter As String 
 
 strScope = "Inbox" 
 
 strFilter = "urn:schemas:httpmail:subject = 'Test'" 
 
 Set sch = Application.AdvancedSearch(strScope, strFilter) 
 
 sch.Stop 
 
End Sub 
 
 
 
Private Sub Application_AdvancedSearchStopped(ByVal SearchObject As Search) 
 
 'Inform the user that the search has stopped. 
 
 MsgBox "An AdvancedSearch has been interrupted and stopped. " 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]