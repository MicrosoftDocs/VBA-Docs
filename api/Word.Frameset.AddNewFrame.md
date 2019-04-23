---
title: Frameset.AddNewFrame method (Word)
keywords: vbawd10.chm165806130
f1_keywords:
- vbawd10.chm165806130
ms.prod: word
api_name:
- Word.Frameset.AddNewFrame
ms.assetid: 81366e66-ae4e-24ce-d7ca-ae6f9273f745
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.AddNewFrame method (Word)

Adds a new frame to a frames page.


## Syntax

_expression_. `AddNewFrame`( `_Where_` )

_expression_ Required. A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Where_|Required| **WdFramesetNewFrameLocation**|Sets the location where the new frame is to be added in relation to the specified frame.|

## Remarks

For more information on creating frames pages, see [Creating frames pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example adds a new frame to the immediate right of the specified frame.


```vb
ActiveDocument.ActiveWindow.ActivePane.Frameset _ 
 .AddNewFrame wdFramesetNewFrameRight
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]