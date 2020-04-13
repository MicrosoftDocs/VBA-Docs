---
title: BuildingBlockEntries object (Word)
keywords: vbawd10.chm553
f1_keywords:
- vbawd10.chm553
ms.prod: word
api_name:
- Word.BuildingBlockEntries
ms.assetid: 9c5946e9-947d-7284-ab16-b570bf7f0ff3
ms.date: 06/08/2017
localization_priority: Normal
---


# BuildingBlockEntries object (Word)

Represents a collection of all  **[BuildingBlock](Word.BuildingBlock.md)** objects in a template.


## Remarks

Use the **[Add](Word.BuildingBlockEntries.Add.md)** method to create a new building block and add it to a template. The following example adds the selected text to the watermarks building block gallery of the first template in the **[Templates](Word.templates.md)** collection.


```vb
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = Templates(1) 
 
Set objBB = objTemplate.BuildingBlockEntries _ 
 .Add(Name:="New Building Block Entry", _ 
 Type:=wdTypeWatermarks, _ 
 Category:="General", _ 
 Range:=Selection.Range)
```

Unlike the **Add** method for the **BuildingBlocks** collection, you need to specify the type and category when you add a building block using the **Add** method of the **BuildingBlockEntries** collection. This is because building blocks are organized by using types and categories. When you use the **BuildingBlockEntries** collection, you are accessing the entire collection of building blocks in a template; however, when you use the **BuildingBlocks** collection, you are accessing the collection of building blocks for a specific type and category in a template.


> [!NOTE] 
> Using the **Category** and **Type** properties for the **BuildingBlock** object enables you to determine the category and type for a building block.

For more information about building blocks, see [Working with Building Blocks](../word/Concepts/Working-with-Word/working-with-building-blocks.md).


## Methods

- [Add](Word.BuildingBlockEntries.Add.md)
- [Item](Word.BuildingBlockEntries.Item.md)

## Properties

- [Application](Word.BuildingBlockEntries.Application.md)
- [Count](Word.BuildingBlockEntries.Count.md)
- [Creator](Word.BuildingBlockEntries.Creator.md)
- [Parent](Word.BuildingBlockEntries.Parent.md)


## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]