---
title: BuildingBlockTypes object (Word)
keywords: vbawd10.chm2896
f1_keywords:
- vbawd10.chm2896
ms.prod: word
api_name:
- Word.BuildingBlockTypes
ms.assetid: fb179437-b736-dd99-3aea-125346aa7a3d
ms.date: 06/08/2017
localization_priority: Normal
---


# BuildingBlockTypes object (Word)

Represents a collection of  **[BuildingBlockType](Word.BuildingBlockType.md)** objects.


## Remarks

Building block types are represented by  **[WdBuildingBlockTypes](Word.WdBuildingBlockTypes.md)** constants. Use the **[Item](Word.BuildingBlockTypes.Item.md)** method to access a specific type in the **BuildingBlockTypes** collection.

To loop through the different building block types, use a  **For** loop with the **[Count](Word.BuildingBlockTypes.Count.md)** property. The following example loops through the building block types and prints the name in the **Immediate Window**. (This example assumes that the **Immediate Window** is visible.)




```vb
Dim objTemplate As Template 
Dim intCount As Integer 
Dim objBBT As BuildingBlockType 
 
Set objTemplate = Templates(1) 
 
For intCount = 1 To objTemplate.BuildingBlockTypes.Count 
 Set objBBT = objTemplate.BuildingBlockTypes(intCount) 
 Debug.Print objBBT.Name 
Next
```

For more information about building blocks, see [Working with Building Blocks](../word/Concepts/Working-with-Word/working-with-building-blocks.md).

## Methods

- [Item](Word.BuildingBlockTypes.Item.md)

## Properties

- [Application](Word.BuildingBlockTypes.Application.md)
- [Count](Word.BuildingBlockTypes.Count.md)
- [Creator](Word.BuildingBlockTypes.Creator.md)
- [Parent](Word.BuildingBlockTypes.Parent.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]