---
title: Subdocuments.AddFromRange method (Word)
keywords: vbawd10.chm159907941
f1_keywords:
- vbawd10.chm159907941
ms.prod: word
api_name:
- Word.Subdocuments.AddFromRange
ms.assetid: ca205880-99d4-2cc5-cb45-3fd8fd60cf36
ms.date: 06/08/2017
localization_priority: Normal
---


# Subdocuments.AddFromRange method (Word)

Creates one or more subdocuments from the text in the specified range and returns a  **SubDocument** object.


## Syntax

_expression_. `AddFromRange`( `_Range_` )

_expression_ Required. A variable that represents a '[Subdocuments](Word.subdocuments.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range used to create one or more subdocuments.|

## Return value

SubDocument


## Remarks

The Range argument must begin with one of the built-in heading level styles (for example, Heading 1). Subdocuments are created at each paragraph formatted with the same heading format used at the beginning of the range. Subdocument files are saved when the master document is saved and are automatically named using text from the first line in the file.


## Example

This example creates one or more subdocuments from the selection.


```vb
ActiveDocument.ActiveWindow.View.Type = wdMasterView 
ActiveDocument.SubDocuments.AddFromRange Range:=Selection.Range
```


## See also


[Subdocuments Collection Object](Word.subdocuments.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]