---
title: Evaluate method (Excel Graph)
keywords: vbagr10.chm65537
f1_keywords:
- vbagr10.chm65537
ms.prod: excel
api_name:
- Excel.Evaluate
ms.assetid: d5f49471-9047-6f72-1f0e-ccd891e73724
ms.date: 04/09/2019
localization_priority: Normal
---


# Evaluate method (Excel Graph)

Converts a Graph name to an object or a value.

## Syntax

_expression_.**Evaluate** (_Name_)

_expression_ Required. An expression that returns an **[Application](excel.application-graph-object.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Name_ |Required |**String**|The name of the specified object, using the Graph naming convention.|

## Remarks

You can use the following types of names in Graph with this method:

- A1-style references. You can use any reference to a single cell in A1-style notation. All references are considered to be absolute references.
    
- Ranges. You can use the range, intersect, and union operators (colon, space, and comma, respectively) with references.
    
- Defined names. You can specify any name in the language of the macro.
    
> [!NOTE] 
> Using square brackets (for example, "[A1:C5]") is identical to calling the **Evaluate** method with a string argument. For example, the following expressions are equivalent.

```vb
myChart.Application.[a1].Value = 25 
myChart.Application.Evaluate("A1").Value = 25
```

The advantage of using square brackets is that the code is shorter. The advantage of using **Evaluate** is that the argument is a string, so you can either construct the string in your code or use a Visual Basic variable.

## Example

This example clears cell A1 on the datasheet.

```vb
clearCell = "A1" 
myChart.Application.Evaluate(clearCell).Clear
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]