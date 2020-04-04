---
title: Explorer.ClearSearch method (Outlook)
keywords: vbaol11.chm2783
f1_keywords:
- vbaol11.chm2783
ms.prod: outlook
api_name:
- Outlook.Explorer.ClearSearch
ms.assetid: 644b6012-0b87-b4cb-6104-6f05b5c4dcc5
ms.date: 06/08/2017
localization_priority: Normal
---


# Explorer.ClearSearch method (Outlook)

Clears results from a Microsoft Instant Search in an **[Explorer](Outlook.Explorer.md)** if results are displayed in the **Explorer**.


## Syntax

_expression_. `ClearSearch`

_expression_ A variable that represents an **[Explorer](Outlook.Explorer.md)** object.


## Remarks

The functionality of this method is analogous to the  **Clear** button in Instant Search.

If no search results are displayed in the Explorer,  **ClearSearch** will not take any action. If the current view of the **Explorer** does not present a search view, **ClearSearch** will not raise an error.


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]