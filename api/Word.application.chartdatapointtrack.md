---
title: Application.ChartDataPointTrack property (Word)
keywords: vbawd10.chm158335470
f1_keywords:
- vbawd10.chm158335470
ms.prod: word
ms.assetid: dea8365d-aadf-6667-ade0-2bef1622fd39
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ChartDataPointTrack property (Word)

Returns or sets a  **Boolean** that specifies whether charts use cell-reference data-point tracking. Read/write.


## Syntax

_expression_. `ChartDataPointTrack`

_expression_ A variable that represents a [Application](./Word.Application.md) object.


## Remarks

In cell-reference data-point tracking, data labels track the cell reference that contains the value of the data point, rather than the index number of the data point. Doing so helps to preserve custom formatting applied by the user, even when a chart is sorted or filtered. Setting  **ChartDataPointTrack** to **True** specifies that charts use cell-reference data-point tracking.


## Property value

 **BOOL**


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]