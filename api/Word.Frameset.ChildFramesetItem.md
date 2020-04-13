---
title: Frameset.ChildFramesetItem property (Word)
keywords: vbawd10.chm165806086
f1_keywords:
- vbawd10.chm165806086
ms.prod: word
api_name:
- Word.Frameset.ChildFramesetItem
ms.assetid: a0210de1-5556-0c20-a694-a6892dc7eddf
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.ChildFramesetItem property (Word)

Returns the **Frameset** object that represents the child **Frameset** object specified by the Index argument. Read-only.


## Syntax

_expression_. `ChildFramesetItem` (_Index_)

 _expression_ An expression that returns a '[Frameset](Word.Frameset.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the specified frame.|

## Remarks

This property applies only to  **Frameset** objects of type **wdFramesetTypeFrameset**.

For more information on creating frames pages, see [Creating Frames Pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example sets the name of the third child frame of the specified frame to "BottomFrame".


```vb
ActiveWindow.Document.Frameset _ 
 .ChildFramesetItem(3).FrameName = "BottomFrame"
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]