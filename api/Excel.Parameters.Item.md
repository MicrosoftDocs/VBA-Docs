---
title: Parameters.Item method (Excel)
keywords: vbaxl10.chm525075
f1_keywords:
- vbaxl10.chm525075
ms.prod: excel
api_name:
- Excel.Parameters.Item
ms.assetid: 66db6a11-b0e3-4417-0589-b0085f67c77a
ms.date: 05/03/2019
localization_priority: Normal
---


# Parameters.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Parameters](Excel.Parameters.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number for the object.|

## Return value

A **[Parameter](Excel.Parameter.md)** object contained by the collection.


## Remarks

The text name of the object is the value of the **[Name](Excel.Parameter.Name.md)** and **[Value](Excel.Parameter.Value.md)** properties.


## Example

This example modifies the parameter prompt string.

```vb
With Worksheets(1).QueryTables(1).Parameters.Item(1) 
 .SetParam xlPrompt, "Please " & .PromptString 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]