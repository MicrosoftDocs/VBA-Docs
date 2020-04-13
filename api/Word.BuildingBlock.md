---
title: BuildingBlock object (Word)
keywords: vbawd10.chm3107
f1_keywords:
- vbawd10.chm3107
ms.prod: word
api_name:
- Word.BuildingBlock
ms.assetid: 2558b89f-8552-bb71-fa40-101cab2635ba
ms.date: 06/08/2017
localization_priority: Normal
---


# BuildingBlock object (Word)

Represents a building block in a template. A building block is pre-built content, similar to autotext, that may contain text, images, and formatting.


## Remarks

Each  **BuildingBlock** object is a member of the **[BuildingBlocks](Word.BuildingBlocks.md)** and **[BuildingBlockEntries](Word.BuildingBlockEntries.md)** collections. Building blocks are stored in Microsoft Word templates. Therefore, to access the building blocks available for a document, you need to access an attached template. Built-in building blocks are stored in the template named "Building Blocks.dotx".

 Use the **[Item](Word.BuildingBlocks.Item.md)** method of the collection or the **BuildingBlocks** collection to return an individual building block. The following example accesses the first building block in the first template in the **[Templates](Word.templates.md)** collection.




```vb
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = Templates(1) 
Set objBB = objTemplate.BuildingBlockEntries.Item(1)
```


> [!NOTE] 
> Depending on how you access the collection, the collection returned may change. For example, if you access a collection of building blocks with a type of  **wdTypeAutoText** with a category of "General", the returned collection may be different from the collection returned if you access a collection of building blocks with a type of **wdTypeAutoText** with a category of "Custom". It is also different from the collection returned if you access the collection of building blocks with a type of **wdTypeCustomAutoText** with a category of "General". Therefore, the first item in a collection accessed from the **BuildingBlockEntries** collection may be different from the first item in the collection accessed from the **BuildingBlocks** collection.

To create a new building block, you can use the **Add** method for either the **BuildingBlockEntries** collection or the **BuildingBlocks** collection. However, the recommended way to create a new building block is by using the **[Add](Word.BuildingBlockEntries.Add.md)** method for the **BuildingBlockEntries** collection. The following example adds the selected text to the watermarks building block gallery of the first template in the **[Templates](Word.templates.md)** collection.




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

Use the **[Insert](Word.BuildingBlock.Insert.md)** method to insert a new building block into a document. The following example inserts the first building block in the first template into the active document at the Insertion Point.




```vb
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = Templates(1) 
Set objBB = objTemplate.BuildingBlockEntries.Item(1) 
 
objBB.Insert Selection.Range
```

Use the **[Delete](Word.BuildingBlock.Delete.md)** method to remove a building block from a template. The following example deletes the first building block from the first template in the **Templates** collection.




```vb
Dim objTemplate As Template 
 
Set objTemplate = Templates(1) 
 
objTemplate.BuildingBlockEntries(1).Delete
```

 Building blocks are organized by category and type. Use the **[BuildingBlockTypes](Word.BuildingBlockTypes.md)** collection to access individual **[BuildingBlockType](Word.BuildingBlockType.md)** objects. Use the **[Categories](Word.Categories.md)** collection to access individual **[Category](Word.BuildingBlock.Category.md)** objects. Then use the **BuildingBlocks** property to access the **BuildingBlocks** collection for a **Category** object. The following example prints the type and category names of all the building blocks in the first template to the **Immediate Window**. (This example assumes that the **Immediate Window** is visible.)




```vb
Dim objTemplate As Template 
Dim objBBT As BuildingBlockType 
Dim objCat As Category 
Dim intCount As Integer 
Dim intCountCat As Integer 
 
Set objTemplate = Templates(1) 
 
For intCount = 1 To objTemplate.BuildingBlockTypes.Count 
 Set objBBT = objTemplate.BuildingBlockTypes(intCount) 
 If objBBT.Categories.Count > 0 Then 
 Debug.Print objBBT.Name 
 For intCountCat = 1 To objBBT.Categories.Count 
 Set objCat = objBBT.Categories(intCountCat) 
 Debug.Print vbTab & objCat.Name 
 Next 
 End If 
Next
```

Each building block has properties that contain information that applies uniquely to it, such as  **[Name](Word.BuildingBlock.Name.md)**, **[Description](Word.BuildingBlock.Description.md)**, **[Type](Word.BuildingBlock.Type.md)**, and **[Value](Word.BuildingBlock.Value.md)**.

For more information about building blocks, see [Working with Building Blocks](../word/Concepts/Working-with-Word/working-with-building-blocks.md).


## Methods

- [Delete](Word.BuildingBlock.Delete.md)
- [Insert](Word.BuildingBlock.Insert.md)

## Properties

- [Application](Word.BuildingBlock.Application.md)
- [Category](Word.BuildingBlock.Category.md)
- [Creator](Word.BuildingBlock.Creator.md)
- [Description](Word.BuildingBlock.Description.md)
- [ID](Word.BuildingBlock.ID.md)
- [Index](Word.BuildingBlock.Index.md)
- [InsertOptions](Word.BuildingBlock.InsertOptions.md)
- [Name](Word.BuildingBlock.Name.md)
- [Parent](Word.BuildingBlock.Parent.md)
- [Type](Word.BuildingBlock.Type.md)
- [Value](Word.BuildingBlock.Value.md)


## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]