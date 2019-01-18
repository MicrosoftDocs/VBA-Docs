---
title: Range.Information property (Word)
keywords: vbawd10.chm157155641
f1_keywords:
- vbawd10.chm157155641
ms.prod: word
api_name:
- Word.Range.Information
ms.assetid: 967e9a22-5f98-e4bd-557c-7367cb7c5d2b
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Information property (Word)

Returns information about the specified range. Read-only **Variant**.

## Syntax

_expression_. `Information`( _Type_ )

_expression_ Required. A variable that represents a [Range](Word.Range.md) object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **wdInformation**|The information type.|

## Example

If the tenth word is in a table, this example selects the table.

```vb
If ActiveDocument.Words(10).Information(wdWithInTable) Then _ 
 ActiveDocument.Words(10).Tables(1).Select
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]