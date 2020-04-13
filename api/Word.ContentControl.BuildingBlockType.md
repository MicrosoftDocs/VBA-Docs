---
title: ContentControl.BuildingBlockType property (Word)
keywords: vbawd10.chm266534933
f1_keywords:
- vbawd10.chm266534933
ms.prod: word
api_name:
- Word.ContentControl.BuildingBlockType
ms.assetid: 6fe28ef5-fe7d-792e-f73a-b6726c802790
ms.date: 06/08/2017
localization_priority: Normal
---


# ContentControl.BuildingBlockType property (Word)

Returns or sets a  **WdBuildingBlockTypes** constant that represents they type of building block for a building block content control. Read/write.


## Syntax

_expression_. `BuildingBlockType`

 _expression_ An expression that returns a [ContentControl](./Word.ContentControl.md) object.


## Remarks

This property applies only to building block content controls and corresponds with the **Gallery** option in the **Content Control Properties** dialog box. You can set this property only for the following building block types:


- Custom 1 through Custom 5
    
- Autotext
    
- Quick Parts
    
- Custom Autotext
    
- Custom Quick Parts
    
- Equations
    

## Example

The following example creates a new building block content control and specifies the type of building block and the gallery.


```vb
Dim objBB As ContentControl 
 
Set objBB = Selection.ContentControls.Add(wdContentControlBuildingBlockGallery) 
 
objBB.BuildingBlockType = wdTypeEquations 
objBB.BuildingBlockCategory = "General"
```


## See also


[ContentControl Object](Word.ContentControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]