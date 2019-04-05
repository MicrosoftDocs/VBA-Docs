---
title: Application.ChartDataPointTrack property (Excel)
keywords: vbaxl10.chm133341
f1_keywords:
- vbaxl10.chm133341
ms.prod: excel
ms.assetid: 124b4d82-de33-c5df-7aa0-1a9c3484a680
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.ChartDataPointTrack property (Excel)

**True** causes all charts in newly created documents to use the cell reference tracking behavior. **Boolean**.


## Syntax

_expression_.**ChartDataPointTrack**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

Data labels now track the _actual_ data point to which they are attached (as opposed to the legacy behavior of tracking the _index_ of the data point), allowing the label-to-point relationship to persist across events such as filter and sort.


## Property value

**BOOL**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]