---
title: Pages.Add method (Outlook)
keywords: vbaol11.chm397
f1_keywords:
- vbaol11.chm397
ms.prod: outlook
api_name:
- Outlook.Pages.Add
ms.assetid: 4a28aac5-be6f-0892-0fc1-17ded4dff783
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.Add method (Outlook)

Creates a new page in the  **[Pages](Outlook.Pages.md)** collection.


## Syntax

_expression_.**Add** `_Name_`

_expression_ A variable that represents a [Pages](Outlook.Pages.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**| The name of the page.|

## Return value

A  **[Page](Outlook.page.md)** object that represents the new page.


## Remarks

The  **Pages** collection is initially empty, and there is a limit of 5 customizable pages per collection.


## See also


[Pages object (Outlook)](Outlook.Pages.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]