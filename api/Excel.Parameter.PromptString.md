---
title: Parameter.PromptString property (Excel)
keywords: vbaxl10.chm523075
f1_keywords:
- vbaxl10.chm523075
ms.prod: excel
api_name:
- Excel.Parameter.PromptString
ms.assetid: e385bffd-fa89-a4c3-6442-d01d957f42d6
ms.date: 05/03/2019
localization_priority: Normal
---


# Parameter.PromptString property (Excel)

Returns the phrase that prompts the user for a parameter value in a parameter query. Read-only **String**.


## Syntax

_expression_.**PromptString**

_expression_ A variable that represents a **[Parameter](Excel.Parameter.md)** object.


## Example

This example modifies the parameter prompt string for query table one.

```vb
With Worksheets(1).QueryTables(1).Parameters(1) 
 .SetParam xlPrompt, "Please " & .PromptString 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]