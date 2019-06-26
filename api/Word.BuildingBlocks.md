---
title: BuildingBlocks object (Word)
ms.prod: word
api_name:
- Word.BuildingBlocks
ms.assetid: be5bba4a-b06c-0074-20bd-bbeb40e03d1c
ms.date: 06/08/2017
localization_priority: Normal
---


# BuildingBlocks object (Word)

Represents a collection of  **[BuildingBlock](Word.BuildingBlock.md)** objects for a specific building block type and category in a template.


## Remarks

Use the  **[Add](Word.BuildingBlocks.Add.md)** method to create a new building block and add it to a template. The following example adds the selected text to the watermarks building block gallery of the first template in the **[Templates](Word.templates.md)** collection.


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

The collection returned with the  **BuildingBlocks** collection is a filtered collection based on the type and category. Depending on how you access the collection, the collection returned changes. For example, if you access a collection of building blocks with a type of **wdTypeAutoText** with a category of "General", the returned collection may be different from the collection returned if you access a collection of building blocks with a type of **wdTypeAutoText** with a category of "Custom". It is also different from the collection returned if you access the collection of building blocks with a type of **wdTypeCustomAutoText** with a category of "General".

For more information about building blocks, see [Working with Building Blocks](../word/Concepts/Working-with-Word/working-with-building-blocks.md).

## Methods

- [Add](Word.BuildingBlocks.Add.md)
- [Item](Word.BuildingBlocks.Item.md)

## Properties

- [Application](Word.BuildingBlocks.Application.md)
- [Count](Word.BuildingBlocks.Count.md)
- [Creator](Word.BuildingBlocks.Creator.md)
- [Parent](Word.BuildingBlocks.Parent.md)


## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]