---
title: ControlFormat.List method (Excel)
keywords: vbaxl10.chm630080
f1_keywords:
- vbaxl10.chm630080
ms.prod: excel
api_name:
- Excel.ControlFormat.List
ms.assetid: 8ec9abd2-d5cf-8179-96e9-a8b583bb8bcc
ms.date: 04/23/2019
localization_priority: Normal
---


# ControlFormat.List method (Excel)

Returns or sets the text entries in the specified list box or combo box, as an array of strings, or returns or sets a single text entry. An error occurs if there are no entries in the list.


## Syntax

_expression_.**List** (_Index_)

_expression_ A variable that represents a **[ControlFormat](Excel.ControlFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The index number of a single text entry to be set or returned. If this argument is omitted, the entire list is returned or set as an array of strings.|

## Return value

Variant


## Remarks

Setting this property clears any range specified by the **[ListFillRange](Excel.ControlFormat.ListFillRange.md)** property.


## Example

This example sets the entries in a list box on worksheet one. If `Shapes(2)` doesn't represent a list box, this example fails.

```vb
Worksheets(1).Shapes(2).ControlFormat.List = _ 
 Array("cogs", "widgets", "sprockets", "gizmos")
```

<br/>

This example sets entry four in a list box on worksheet one. If `Shapes(2)` doesn't represent a list box, this example fails.

```vb
Worksheets(1).Shapes(2).ControlFormat.List(4) = "gadgets"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]