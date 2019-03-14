---
title: Form.VerticalDatasheetGridlineStyle property (Access)
keywords: vbaac10.chm13515
f1_keywords:
- vbaac10.chm13515
ms.prod: access
api_name:
- Access.Form.VerticalDatasheetGridlineStyle
ms.assetid: b0174311-f03b-aa6a-b15a-697f6be1b2ac
ms.date: 03/15/2019
localization_priority: Normal
---


# Form.VerticalDatasheetGridlineStyle property (Access)

Returns or sets a **Byte** indicating the line style to use for vertical gridlines on the specified datasheet. Read/write.


## Syntax

_expression_.**VerticalDatasheetGridlineStyle**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values are between zero and seven. Values greater than seven are ignored; negative values or values above 255 cause an error.

|Value|Description|
|:-----|:-----|
|0|Transparent border|
|1|Solid|
|2|Dashes|
|3|Short dashes|
|4|Dots|
|5|Sparse dots|
|6|Dash-dot|
|7|Dash-dot-dot|

## Example

This example sets the vertical gridline style on the first open form to dashes. The form must be set to Datasheet view for you to see the change.

```vb
Forms(0).VerticalDatasheetGridlineStyle = 2 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]