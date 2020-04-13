---
title: TimeZone.ID property (Outlook)
keywords: vbaol11.chm3304
f1_keywords:
- vbaol11.chm3304
ms.prod: outlook
api_name:
- Outlook.TimeZone.ID
ms.assetid: 13d4826f-5291-993c-2da1-f1dc65a1e086
ms.date: 06/08/2017
localization_priority: Normal
---


# TimeZone.ID property (Outlook)

Returns a **String** that uniquely identifies the time zone. Read-only.


## Syntax

_expression_.**ID**

_expression_ A variable that represents a [TimeZone](Outlook.TimeZone.md) object.


## Remarks

The **ID** of a time zone is globally the same for that time zone. It is the name of the Windows registry key that contains the time zone information. Unlike the **[Name](Outlook.TimeZone.Name.md)** property, the value of **ID** is not localized.


## See also


[TimeZone Object](Outlook.TimeZone.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]