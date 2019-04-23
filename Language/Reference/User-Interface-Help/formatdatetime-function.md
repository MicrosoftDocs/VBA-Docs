---
title: FormatDateTime function (Visual Basic for Applications)
keywords: vblr6.chm1011367
f1_keywords:
- vblr6.chm1011367
ms.prod: office
ms.assetid: 1ead64ea-cea4-0464-a6e4-f28b1edb06cc
ms.date: 12/12/2018
localization_priority: Normal
---


# FormatDateTime function

Returns an expression formatted as a date or time.

## Syntax

**FormatDateTime**(_Date_, [ _NamedFormat_ ])

<br/>

The **FormatDateTime** function syntax has these parts:

|Part|Description|
|:-----|:-----|
| _Date_|Required. Date expression to be formatted.|
| _NamedFormat_|Optional. Numeric value that indicates the date/time format used. If omitted, **vbGeneralDate** is used.|

## Settings

The _NamedFormat_ argument has the following settings:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbGeneralDate**|0|Display a date and/or time. If there is a date part, display it as a short date. If there is a time part, display it as a long time. If present, both parts are displayed.|
|**vbLongDate**|1|Display a date by using the long date format specified in your computer's regional settings.|
|**vbShortDate**|2|Display a date by using the short date format specified in your computer's regional settings.|
|**vbLongTime**|3|Display a time by using the time format specified in your computer's regional settings.|
|**vbShortTime**|4|Display a time by using the 24-hour format (hh:mm).|

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
