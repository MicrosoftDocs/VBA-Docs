---
title: BuildingBlock.Insert method (Word)
keywords: vbawd10.chm203620454
f1_keywords:
- vbawd10.chm203620454
ms.prod: word
api_name:
- Word.BuildingBlock.Insert
ms.assetid: e2f3fd61-624b-fd18-3b5a-2c9f16fa6bd2
ms.date: 06/08/2017
localization_priority: Normal
---


# BuildingBlock.Insert method (Word)

Inserts the value of a building block into a document and returns a  **[Range](Word.Range.md)** object that represents the contents of the building block within the document.


## Syntax

_expression_.**Insert** (_Where_, _RichText_)

 _expression_ An expression that returns a '[BuildingBlock](Word.BuildingBlock.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Where_|Required| **Range**|The location of where to place the contents of the building block.|
| _RichText_|Optional| **Variant**| **True** inserts the building block as rich, formatted text. **False** inserts the building block as plain text.|

## Return value

Range


## Example

The following example inserts the first building block from the first template into the first paragraph of the active document.


```vb
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = Templates(1) 
Set objBB = objTemplate.BuildingBlockEntries(1) 
 
objBB.Insert ActiveDocument.Paragraphs(1).Range
```


## See also


[BuildingBlock Object](Word.BuildingBlock.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]