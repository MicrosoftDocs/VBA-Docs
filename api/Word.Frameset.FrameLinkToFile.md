---
title: Frameset.FrameLinkToFile property (Word)
keywords: vbawd10.chm165806117
f1_keywords:
- vbawd10.chm165806117
ms.prod: word
api_name:
- Word.Frameset.FrameLinkToFile
ms.assetid: a27ce637-a892-3697-a727-e7c60eb26aaf
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset.FrameLinkToFile property (Word)

 **True** if the webpage or other document specified by the **[FrameDefaultURL](Word.Frameset.FrameDefaultURL.md)** property is an external file to which Microsoft Word maintains only a link from the specified frame. Read/write **Boolean**.


## Syntax

_expression_. `FrameLinkToFile`

_expression_ A variable that represents a '[Frameset](Word.Frameset.md)' object.


## Remarks

For more information on creating frames pages, see [Creating Frames Pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## Example

This example sets Microsoft Word to maintain only a link from the specified frame to the document "Order.htm".


```vb
With ActiveDocument.ActiveWindow.ActivePane.Frameset 
 .FrameDefaultURL = "C:\Documents\Order.htm" 
 .FrameLinkToFile = True 
End With
```


## See also


[Frameset Object](Word.Frameset.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]