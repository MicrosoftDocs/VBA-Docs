---
title: Categories object (Word)
ms.prod: word
api_name:
- Word.Categories
ms.assetid: f5f5081d-4309-6617-28da-c369c1fe690c
ms.date: 06/08/2017
localization_priority: Normal
---


# Categories object (Word)

Represents a collection of building block categories.


## Remarks

Use the **Item** method to access an existing category. You can then use the **[BuildingBlocks](Word.Category.BuildingBlocks.md)** property to access a collection of **[BuildingBlock](Word.BuildingBlock.md)** objects for the category. The following example prints the type and category names of all the building blocks in the first template to the **Immediate Window**. (This example assumes that the **Immediate Window** is visible.)


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

Use the **Item** method to access an existing category; to create a new category, use the **Add** method of the **BuildingBlockEntries** collection. Set the value of the Category parameter.

For more information about building blocks, see [Working with Building Blocks](../word/Concepts/Working-with-Word/working-with-building-blocks.md).


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]