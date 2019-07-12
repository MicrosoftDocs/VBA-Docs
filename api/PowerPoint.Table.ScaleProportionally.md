---
title: Table.ScaleProportionally method (PowerPoint)
keywords: vbapp10.chm622016
f1_keywords:
- vbapp10.chm622016
ms.prod: powerpoint
api_name:
- PowerPoint.Table.ScaleProportionally
ms.assetid: 1c703fe7-d657-5588-1991-23304a5b2bda
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.ScaleProportionally method (PowerPoint)

Scales all cell heights and widths, font sizes, and internal margins in the table by a specified proportion.


## Syntax

_expression_. `ScaleProportionally`( `_scale_` )

_expression_ A variable that represents a [Table](PowerPoint.Table.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _scale_|Required|**Single**|The proportion to scale the table, between 0.01 and 100. For example, a scale value of 1 keeps the table layout unchanged; a value of 2 makes it twice as large; a value of 0.5 makes it half the size.|

## Remarks

Use the  **ScaleProportionally** method to resize a table and maintain the text layout as close as possible to the original layout.


## See also


[Table Object](PowerPoint.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]