---
title: Form.DatasheetColumnHeaderUnderlineStyle property (Access)
keywords: vbaac10.chm13513
f1_keywords:
- vbaac10.chm13513
ms.prod: access
api_name:
- Access.Form.DatasheetColumnHeaderUnderlineStyle
ms.assetid: 9e689097-f3ed-bcda-9cc5-d423a3b92806
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.DatasheetColumnHeaderUnderlineStyle property (Access)

Returns or sets a **Byte** indicating the line style to use for the bottom edge of the column headers on the specified datasheet. Read/write.


## Syntax

_expression_.**DatasheetColumnHeaderUnderlineStyle**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

Valid values are between zero and eight. Values greater than eight are ignored; negative values or values above 255 cause an error.

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
|8|Double solid|

## Example

This example sets the column header underline style for the first open form to sparse dots. The form must be set to Datasheet view for you to see the change.

```vb
Forms(0).DatasheetColumnHeaderUnderlineStyle = 5 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]