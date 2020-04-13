---
title: DropCap.Position property (Word)
keywords: vbawd10.chm156631050
f1_keywords:
- vbawd10.chm156631050
ms.prod: word
api_name:
- Word.DropCap.Position
ms.assetid: ab217570-e506-6fd5-6e8d-4321925907ee
ms.date: 06/08/2017
localization_priority: Normal
---


# DropCap.Position property (Word)

Returns or sets the position of a dropped capital letter. Read/write  **WdDropPosition**.


## Syntax

_expression_.**Position**

_expression_ Required. A variable that represents a '[DropCap](Word.DropCap.md)' object.


## Example

This example sets the first paragraph in the active document to begin with a dropped capital letter. The position of the **DropCap** object is set to **wdDropNormal**.


```vb
With ActiveDocument.Paragraphs(1).DropCap 
 .Enable 
 .FontName= "Arial" 
 .Position = wdDropNormal 
End With
```


## See also


[DropCap Object](Word.DropCap.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]