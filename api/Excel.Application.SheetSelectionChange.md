---
title: Application.SheetSelectionChange event (Excel)
keywords: vbaxl10.chm504074
f1_keywords:
- vbaxl10.chm504074
ms.prod: excel
api_name:
- Excel.Application.SheetSelectionChange
ms.assetid: c98203c2-b306-d8b7-b75f-1304be7b5751
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.SheetSelectionChange event (Excel)

Occurs when the selection changes on any worksheet (doesn't occur if the selection is on a chart sheet).


## Syntax

_expression_.**SheetSelectionChange** (_Sh_, _Target_)

_expression_ An expression that returns an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The worksheet that contains the new selection.|
| _Target_|Required| **Range**|The new selected range.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]