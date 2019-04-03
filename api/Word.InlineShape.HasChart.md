---
title: InlineShape.HasChart property (Word)
keywords: vbawd10.chm162005140
f1_keywords:
- vbawd10.chm162005140
ms.prod: word
api_name:
- Word.InlineShape.HasChart
ms.assetid: f8b88eef-ec41-fc03-f58b-e346d240a121
ms.date: 06/08/2017
localization_priority: Normal
---


# InlineShape.HasChart property (Word)

 **True** if the specified shape is a chart. Read-only.


## Syntax

_expression_. `HasChart`

 _expression_ An expression that returns an [InlineShape](./Word.InlineShape.md) object.


## Remarks

This property always returns false for OLE charts. For OLE charts, use  `InlineShape.OLEFormat.ProgID` and check for the following possible values: "Excel.Chart.8", "MSGraph.Chart.8", "Excel.Sheet.8", "Excel.Chart.5", "MSGraph.Chart.5", or "Excel.Sheet.5".


## See also


[InlineShape Object](Word.InlineShape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]