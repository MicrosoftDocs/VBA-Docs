---
title: Set the Building Blocks That You Can Use for a Content Control
ms.prod: word
ms.assetid: 6723a4c4-f96c-7bbd-a978-66602ab693c7
ms.date: 06/08/2019
localization_priority: Normal
---


# Set the Building Blocks That You Can Use for a Content Control

A document building block is a predesigned piece of content, such as a cover page or a header or footer. Word includes a library of document building blocks that users can choose from to insert into a document. 

A [ContentControl object (Word)](../../../api/Word.ContentControl.md) object with a [ContentControl.Type property (Word)](../../../api/Word.ContentControl.Type.md) property value of **wdContentControlBuildingBlockGallery** specifies a content control that can contain document building blocks.

The **[WdBuildingBlockTypes](../../../api/Word.WdBuildingBlockTypes.md)** enumeration contains each building block type. You can only use the following building block types within a building block gallery content control:


- AutoText
    
- Tables
    
- Equations
    
- Quick Parts
    
- Custom 1 though Custom 5
    
- Custom AutoText
    
- Custom Tables
    
- Custom Equations
    
- Custom Quick Parts
    
For more information about content controls, see [Working with Content Controls](../Working-with-Word/working-with-content-controls.md).
The objects used in this sample are:

- **[ContentControl](../../../api/Word.ContentControl.md)**
    
- **[ContentControls](../../../api/Word.ContentControls.md)**
    

## Sample

The following code sample instantiates a building block gallery content control and then adds a building block to the content control.


```vb
Sub SetBuildingBlock()
 
    Dim strTitle As String
    strTitle = "My Equation"
    Dim objContentControl As ContentControl
     
    Set objContentControl = ActiveDocument.ContentControls _
        .Add(wdContentControlBuildingBlockGallery)
    objContentControl.Title = strTitle
    objContentControl.BuildingBlockType = wdTypeEquations
   
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]