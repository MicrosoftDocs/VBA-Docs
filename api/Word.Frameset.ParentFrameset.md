---
title: Frameset.ParentFrameset property (Word)
keywords: vbawd10.chm165807083
f1_keywords:
- vbawd10.chm165807083
ms.prod: word
api_name:
- Word.Frameset.ParentFrameset
ms.assetid: aa2759c6-4072-00c6-0c4f-ef12ecc19bd6
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.ParentFrameset property (Word)

Returns a  **Frameset** object that represents the parent of the specified **Frameset** object on a frames page.


## Syntax

_expression_. `ParentFrameset`

 _expression_ An expression that returns a '[Frameset](Word.Frameset.md)' object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example returns the number of child  **Frameset** objects belonging to the parent **Frameset** object of the specified frame.


```vb
MsgBox ActiveDocument.ActiveWindow.ActivePane _ 
 .Frameset.ParentFrameset.ChildFramesetCount
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]