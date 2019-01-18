---
title: Font2.Fill property (Office)
ms.prod: office
api_name:
- Office.Font2.Fill
ms.assetid: b8f19a98-4e22-d2ad-1404-3ee48d3edde3
ms.date: 01/09/2019
localization_priority: Normal
---


# Font2.Fill property (Office)

Gets the formatting properties for the font of the specified text. Read-only.


## Syntax

_expression_.**Fill**

_expression_ An expression that returns a **[Font2](Office.Font2.md)** object.


## Example

The following code assumes that a shape has been inserted into Sheet1. The code inserts text into the shape and changes the fore color of the font to bold and red. It then adds a carriage return after the second word, creating a second paragraph, and aligns the paragraph to the right.


```vb
Sub TestShapeAttributes() 
 Dim shp As Excel.Shape 
 Dim rng As Office.TextRange2 
 Dim rngWord As Office.TextRange2 
 Dim rngRun As Office.TextRange2 
 Dim rngPara As Office.TextRange2 
 Dim fnt As Office.Font2 
 
 Set shp = ActiveSheet.Shapes(1) 
 Set rng = shp.TextFrame2.TextRange 
 rng.Text = "This is test text." 
 
 Set rngWord = rng.Words(2) 
 Set fnt = rngWord.Font 
 With fnt 
 .Fill.ForeColor.RGB = RGB(255, 0, 0) 
 .Bold = msoTrue 
 End With 
 
 Set rngRun = rng.Runs(3) 
 rngRun.InsertBefore vbCr 
 
 Set rngPara = rng.Paragraphs(2) 
 rngPara.ParagraphFormat.Alignment = msoAlignRight 
End Sub 

```


## See also

- [Font2 object members](overview/library-reference/font2-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]