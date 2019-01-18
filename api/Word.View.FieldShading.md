---
title: View.FieldShading property (Word)
keywords: vbawd10.chm161808407
f1_keywords:
- vbawd10.chm161808407
ms.prod: word
api_name:
- Word.View.FieldShading
ms.assetid: 4e699444-0946-5d58-cf87-456b4bf49be5
ms.date: 06/08/2017
localization_priority: Normal
---


# View.FieldShading property (Word)

Returns or sets on-screen shading for fields. Read/write  **WdFieldShading**.


## Syntax

 _expression_. `FieldShading`

 _expression_ Required. A variable that represents a '[View](Word.View.md)' object.


## Example

This example enables field shading for all form fields in the active window.


```vb
ActiveDocument.ActiveWindow.View.FieldShading = _ 
 wdFieldShadingAlways
```


## See also


[View Object](Word.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]