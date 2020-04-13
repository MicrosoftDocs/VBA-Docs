---
title: TimelineView.Copy method (Outlook)
keywords: vbaol11.chm2647
f1_keywords:
- vbaol11.chm2647
ms.prod: outlook
api_name:
- Outlook.TimelineView.Copy
ms.assetid: 0fb16952-06bb-d8ca-a8f2-9cb2e99fa299
ms.date: 06/08/2017
localization_priority: Normal
---


# TimelineView.Copy method (Outlook)

Creates a new **[View](Outlook.View.md)** object based on the existing **[TimelineView](Outlook.TimelineView.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _SaveOption_)

_expression_ A variable that represents a [TimelineView](Outlook.TimelineView.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new view.|
| _SaveOption_|Optional| **[OlViewSaveOption](Outlook.OlViewSaveOption.md)**|The save option for the new view.|

## Return value

A  **View** object that represents the new view.


## See also


[TimelineView Object](Outlook.TimelineView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]