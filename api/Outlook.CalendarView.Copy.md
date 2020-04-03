---
title: CalendarView.Copy method (Outlook)
keywords: vbaol11.chm2612
f1_keywords:
- vbaol11.chm2612
ms.prod: outlook
api_name:
- Outlook.CalendarView.Copy
ms.assetid: ed33fd43-f36a-99e2-db61-9482423a9558
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarView.Copy method (Outlook)

Creates a new  **[View](Outlook.View.md)** object based on the existing **[CalendarView](Outlook.CalendarView.md)** object.


## Syntax

_expression_.**Copy** (_Name_, _SaveOption_)

_expression_ A variable that represents a [CalendarView](Outlook.CalendarView.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new view.|
| _SaveOption_|Optional| **[OlViewSaveOption](Outlook.OlViewSaveOption.md)**|The save option for the new view.|

## Return value

A  **View** object that represents the new view.


## See also


[CalendarView Object](Outlook.CalendarView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]