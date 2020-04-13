---
title: Category object (Word)
keywords: vbawd10.chm2910
f1_keywords:
- vbawd10.chm2910
ms.prod: word
api_name:
- Word.Category
ms.assetid: 5485ae39-fbcf-b18f-b1f9-945e220ecd2a
ms.date: 06/08/2017
localization_priority: Normal
---


# Category object (Word)

Represents an individual category of a building block type.


## Remarks

Microsoft Word uses types and categories to organize building blocks. Each building block type is represented by a  **[WdBuildingBlockTypes](Word.WdBuildingBlockTypes.md)** constant. Each category is a unique string that a user defines. Word comes with two categories already defined: "General" and "Custom"; you can create additional categories as you need.

Use the **[Type](Word.Category.Type.md)** property to access the building block type associated with a specific category. Use the **[BuildingBlocks](Word.Category.BuildingBlocks.md)** property to access the collection of building blocks for a category. The following example prints the type and category names of all the building blocks in the first template to the **Immediate Window**. (This example assumes that the **Immediate Window** is visible.)




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

Use the **Item** method of the **Categories** collection to access an existing category; to create a new category, use the **Add** method of the **BuildingBlockEntries** collection. Set the value of the Category parameter.

For more information about building blocks, see [Working with Building Blocks](../word/Concepts/Working-with-Word/working-with-building-blocks.md).


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]