---
title: BuildingBlockEntries.Add method (Word)
keywords: vbawd10.chm36241509
f1_keywords:
- vbawd10.chm36241509
ms.prod: word
api_name:
- Word.BuildingBlockEntries.Add
ms.assetid: 09578906-ea6d-9475-e026-b9dc437f451b
ms.date: 06/08/2017
localization_priority: Normal
---


# BuildingBlockEntries.Add method (Word)

Creates a new building block entry in a template and returns a  **[BuildingBlock](Word.BuildingBlock.md)** object that represents the new building block entry.


## Syntax

_expression_.**Add** (_Name_, _Type_, _Category_, _Range_, _Description_, _InsertOptions_)

 _expression_ An expression that returns a '[BuildingBlockEntries](Word.BuildingBlockEntries.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Specifies the name of the building block entry. Corresponds to the **[Name](Word.BuildingBlock.Name.md)** property of the **BuildingBlock** object.|
| _Type_|Required| **[WdBuildingBlockTypes](Word.WdBuildingBlockTypes.md)**|Specifies the type of building block to create. Corresponds to the **[Type](Word.BuildingBlock.Type.md)** property of the **BuildingBlock** object.|
| _Category_|Required| **String**|Specifies the category of the new building block entry. Corresponds to the **[Category](Word.BuildingBlock.Category.md)** property of the **BuildingBlock** object.|
| _Range_|Required| **[Range](Word.Range.md)**|Specifies the value of the buildling block entry. Corresponds to the **[Value](Word.BuildingBlock.Value.md)** property of the **BuildingBlock** object.|
| _Description_|Optional| **Variant**|Specifies the description of the buildling block entry. Corresponds to the **[Description](Word.BuildingBlock.Description.md)** property of the **BuildingBlock** object.|
| _InsertOptions_|Optional| **[WdDocPartInsertOptions](Word.WdDocPartInsertOptions.md)**|Specifies whether the building block entry is inserted as a page, a paragraph, or inline. If omitted, the default value is **wdInsertContent**. Corresponds to the **[InsertOptions](Word.BuildingBlock.InsertOptions.md)** property for the **BuildingBlock** object.|

## Return value

BuildingBlock


## Example

The following example creates a new building block entry and adds it to the template attached to the active document, and than sets the value of the building block to the selected text.


```vb
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = ActiveDocument.AttachedTemplate 
Set objBB = objTemplate.BuildingBlockEntries.Add("Author Name", _ 
 wdTypeCustomTextBox, "Custom", Selection.Range)
```


## See also


[BuildingBlockEntries Collection](Word.BuildingBlockEntries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]