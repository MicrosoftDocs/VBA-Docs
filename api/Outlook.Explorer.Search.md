---
title: Explorer.Search method (Outlook)
keywords: vbaol11.chm2784
f1_keywords:
- vbaol11.chm2784
ms.prod: outlook
api_name:
- Outlook.Explorer.Search
ms.assetid: d4dc7ae5-c24f-90df-f52e-e0b73293e25d
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.Search method (Outlook)

Performs a Microsoft Instant Search on the current folder displayed in the Explorer using the given  _Query_.


## Syntax

_expression_. `Search`( `_Query_` , `_SearchScope_` )

_expression_ A variable that represents an **[Explorer](Outlook.Explorer.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Query_|Required| **String**|A search string that can contain any valid keywords supported in Instant Search.|
| _SearchScope_|Optional| **[OlSearchScope](Outlook.OlSearchScope.md)**|Specifies the scope in terms of folders for the search.|

## Remarks

The functionality of  **Explorer.Search** is analogous to the **Search** button in Instant Search. It behaves as if the user has typed the query string in the Instant Search user interface and then clicked **Search**. When calling  **Search**, the query is run in the user interface, and there is no programmatic mechanism to obtain the search results. For more information on Instant Search, query for "Instant Search" in the Outlook Help.

The **Search** method does not provide a callback to enable the developer to determine when the search is complete.


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]