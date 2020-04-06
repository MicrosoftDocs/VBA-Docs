---
title: Search.IsSynchronous property (Outlook)
keywords: vbaol11.chm2254
f1_keywords:
- vbaol11.chm2254
ms.prod: outlook
api_name:
- Outlook.Search.IsSynchronous
ms.assetid: e240cc55-26c3-a560-4ee2-84b15da95e52
ms.date: 06/08/2017
localization_priority: Normal
---


# Search.IsSynchronous property (Outlook)

Returns a **Boolean** indicating whether the search is synchronous. Read-only.


## Syntax

_expression_. `IsSynchronous`

_expression_ A variable that represents a [Search](Outlook.Search.md) object.


## Remarks

A search can be synchronous or asynchronous. If the search is synchronous, code execution will pause until the search has completed. Conversely, if the search is asynchronous, code execution will continue even though the search has not completed. In this case, use the  **[Search](Outlook.Search.md)** object's **[Stop](Outlook.Search.Stop.md)** method to halt the search. In order to get meaningful results from an asynchronous search, use the **[AdvancedSearchComplete](Outlook.Application.AdvancedSearchComplete.md)** event to notify you when the search has finished.


## See also


[Search Object](Outlook.Search.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]