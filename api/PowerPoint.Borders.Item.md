---
title: Borders.Item method (PowerPoint)
keywords: vbapp10.chm629003
f1_keywords:
- vbapp10.chm629003
ms.prod: powerpoint
api_name:
- PowerPoint.Borders.Item
ms.assetid: fad023e2-55c1-4115-fc61-cd4519486fad
ms.date: 06/08/2017
localization_priority: Normal
---


# Borders.Item method (PowerPoint)

Returns a  **[LineFormat](PowerPoint.LineFormat.md)** object for the specified border from the **Borders** collection.


## Syntax

_expression_.**Item** (_BorderType_)

_expression_ A variable that represents a [Borders](PowerPoint.Borders.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BorderType_|Required|**PpBorderType**|Specifies which border of a cell or cell range is to be returned.|

## Return value

LineFormat


## Remarks

The  _BorderType_ parameter value can be one of these **PpBorderType** constants.


||
|:-----|
|**ppBorderBottom**|
|**ppBorderDiagonalDown**|
|**ppBorderDiagonalUp**|
|**ppBorderLeft**|
|**ppBorderRight**|
|**ppBorderTop**|

## See also


[Borders Object](PowerPoint.Borders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]