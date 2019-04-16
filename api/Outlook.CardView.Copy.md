---
title: CardView.Copy method (Outlook)
keywords: vbaol11.chm2584
f1_keywords:
- vbaol11.chm2584
ms.prod: outlook
api_name:
- Outlook.CardView.Copy
ms.assetid: 36f59955-3bbb-99b4-af1a-3b0165470a89
ms.date: 06/08/2017
localization_priority: Normal
---


# CardView.Copy method (Outlook)

Creates a new  **[View](Outlook.View.md)** object based on the existing **[CardView](Outlook.CardView.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _SaveOption_)

_expression_ A variable that represents a [CardView](Outlook.CardView.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new view.|
| _SaveOption_|Optional| **[OlViewSaveOption](Outlook.OlViewSaveOption.md)**|The save option for the new view.|

## Return value

A  **View** object that represents the new view.


## See also


[CardView Object](Outlook.CardView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]