---
title: CalendarSharing.ForwardAsICal method (Outlook)
keywords: vbaol11.chm2412
f1_keywords:
- vbaol11.chm2412
ms.prod: outlook
api_name:
- Outlook.CalendarSharing.ForwardAsICal
ms.assetid: b796a573-784b-6725-535e-fd156a3f233c
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarSharing.ForwardAsICal method (Outlook)

Forwards calendar information from the parent  **[Folder](Outlook.Folder.md)** of the **[CalendarSharing](Outlook.CalendarSharing.md)** object as the payload of a **[MailItem](Outlook.MailItem.md)**.


## Syntax

_expression_. `ForwardAsICal`( `_MailFormat_` )

 _expression_ An expression that returns a [CalendarSharing](Outlook.CalendarSharing.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MailFormat_|Required| **[OlCalendarMailFormat](Outlook.OlCalendarMailFormat.md)**|Determines the format of the calendar information in the body of the  **MailItem** created by this method.|

## Return value

A **MailItem** object that represents the new mail item to which the calendar information is attached.


## Remarks

The **ForwardAsICal** method provides a single method by which you can use payload sharing to share a calendar with other users. The method:


- Creates a **MailItem** object and provides a presentation of calendar information in the body of the mail item.
    
- Creates an iCalendar (.ics) file containing the calendar information and attaches the file to the  **MailItem**.
    

## See also


[CalendarSharing Object](Outlook.CalendarSharing.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]