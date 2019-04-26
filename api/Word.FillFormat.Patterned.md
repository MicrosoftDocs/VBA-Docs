---
title: FillFormat.Patterned method (Word)
keywords: vbawd10.chm164102156
f1_keywords:
- vbawd10.chm164102156
ms.prod: word
api_name:
- Word.FillFormat.Patterned
ms.assetid: 993fd302-0ba2-f540-f21c-0915bccfacaf
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.Patterned method (Word)

Sets the specified fill to a pattern.


## Syntax

_expression_.**Patterned** (_Pattern_)

_expression_ Required. A variable that represents a **[FillFormat](word.fillformat.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pattern_|Required| **MsoPatternType**|The pattern to be used for the specified fill.|

## Remarks

Use the  **BackColor** and **ForeColor** properties to set the colors used in the pattern.


## Example

This example adds an oval with a patterned fill to the active document.


```vb
Sub FillPattern() 
 With ActiveDocument.Shapes.AddShape _ 
 (msoShapeOval, 60, 60, 80, 40).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(0, 0, 255) 
 .Patterned msoPatternDarkVertical 
 End With 
End Sub
```


## See also


[FillFormat Object](Word.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]