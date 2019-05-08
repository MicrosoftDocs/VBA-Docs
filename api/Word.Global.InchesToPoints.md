---
title: Global.InchesToPoints method (Word)
keywords: vbawd10.chm163119474
f1_keywords:
- vbawd10.chm163119474
ms.prod: word
api_name:
- Word.Global.InchesToPoints
ms.assetid: 7e8f5631-fa6a-702a-5785-da7b34495a22
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.InchesToPoints method (Word)

Converts a measurement from inches to points (1 inch = 72 points). Returns the converted measurement as a  **Single**.


## Syntax

_expression_. `InchesToPoints`( `_Inches_` )

_expression_ A variable that represents a '[Global](Word.Global.md)' object. Optional.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Inches_|Required| **Single**|The inch value to be converted to points.|

## Return value

Single


## Example

This example sets the space before for the selected paragraphs to 0.25 inch.


```vb
Selection.ParagraphFormat.SpaceBefore = InchesToPoints(0.25)
```

This example prints each open document after setting the left and right margins to 0.65 inch.




```vb
Dim docLoop As Document 
 
For Each docLoop in Documents 
 With docLoop 
 .PageSetup.LeftMargin = InchesToPoints(0.65) 
 .PageSetup.RightMargin = InchesToPoints(0.65) 
 .PrintOut 
 End With 
Next docLoop
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]