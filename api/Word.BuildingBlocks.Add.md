---
title: BuildingBlocks.Add method (Word)
ms.prod: word
api_name:
- Word.BuildingBlocks.Add
ms.assetid: 22725f33-4de0-95cd-d4a5-a2379b0130c4
ms.date: 06/08/2017
localization_priority: Normal
---


# BuildingBlocks.Add method (Word)

Creates a new building block and returns a  **BuildingBlock** object.


## Syntax

_expression_.**Add** (_Name_, _Range_, _Description_, _InsertOptions_)

 _expression_ An expression that returns a [BuildingBlocks](./Word.BuildingBlocks.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Specifies the name of the building block entry. Corresponds to the  **[Name](Word.BuildingBlock.Name.md)** property of the **BuildingBlock** object.|
| _Range_|Required| **Range**|Specifies the value of the buildling block entry. Corresponds to the  **[Value](Word.BuildingBlock.Value.md)** property of the **BuildingBlock** object.|
| _Description_|Optional| **Variant**|Specifies the description of the buildling block entry. Corresponds to the  **[Description](Word.BuildingBlock.Description.md)** property of the **BuildingBlock** object.|
| _InsertOptions_|Optional| **[WdDocPartInsertOptions](Word.WdDocPartInsertOptions.md)**|Specifies whether the building block entry is inserted as a page, a paragraph, or inline. If omitted, the default value is  **wdInsertContent**. Corresponds to the **[InsertOptions](Word.BuildingBlock.InsertOptions.md)** property for the **BuildingBlock** object.|

## Return value

BuildingBlock


## Example

The following example adds a new building block auto text entry to the first template in the collection of templates.


```vb
Dim objTemplate As Template 
 
Set objTemplate = Templates(1) 
 
objTemplate.BuildingBlockTypes(wdTypeAutoText) _ 
 .Categories("General").BuildingBlocks _ 
 .Add Name:="New Building Block", _ 
 Range:=Selection.Range
```


## See also


[BuildingBlocks Collection](Word.BuildingBlocks.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]