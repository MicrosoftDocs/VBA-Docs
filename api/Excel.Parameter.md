---
title: Parameter object (Excel)
keywords: vbaxl10.chm522072
f1_keywords:
- vbaxl10.chm522072
ms.prod: excel
api_name:
- Excel.Parameter
ms.assetid: 2a30f4ef-2cae-c96d-4480-3ba55fa871e8
ms.date: 03/30/2019
localization_priority: Normal
---


# Parameter object (Excel)

Represents a single parameter used in a parameter query.


## Remarks

The **Parameter** object is a member of the **[Parameters](Excel.Parameters.md)** collection.


## Example

Use **[Parameters](Excel.QueryTable.Parameters.md)** (_index_), where _index_ is the index number of the parameter, to return a single **Parameter** object. The following example modifies the prompt string for parameter one.

```vb
With Worksheets(1).QueryTables(1).Parameters(1) 
 .SetParam xlPrompt, "Please " & .PromptString 
End With
```

## Methods

- [SetParam](Excel.Parameter.SetParam.md)

## Properties

- [Application](Excel.Parameter.Application.md)
- [Creator](Excel.Parameter.Creator.md)
- [DataType](Excel.Parameter.DataType.md)
- [Name](Excel.Parameter.Name.md)
- [Parent](Excel.Parameter.Parent.md)
- [PromptString](Excel.Parameter.PromptString.md)
- [RefreshOnChange](Excel.Parameter.RefreshOnChange.md)
- [SourceRange](Excel.Parameter.SourceRange.md)
- [Type](Excel.Parameter.Type.md)
- [Value](Excel.Parameter.Value.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]