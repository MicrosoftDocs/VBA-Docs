---
title: DropCap.ApplyCustomDropCap method (Publisher)
keywords: vbapb10.chm5505041
f1_keywords:
- vbapb10.chm5505041
ms.prod: publisher
api_name:
- Publisher.DropCap.ApplyCustomDropCap
ms.assetid: 906cf476-3826-8510-315f-425f6f50a92a
ms.date: 06/07/2019
localization_priority: Normal
---


# DropCap.ApplyCustomDropCap method (Publisher)

Applies custom formatting to the first letters of paragraphs in a text frame.


## Syntax

_expression_.**ApplyCustomDropCap** (_LinesUp_, _Size_, _Span_, _FontName_, _Bold_, _Italic_)

_expression_ A variable that represents a **[DropCap](Publisher.DropCap.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_LinesUp_|Optional| **Long**|The number of lines to move up the dropped capital letter. The default is 0. The maximum number cannot be more than the number entered for the _Size_ argument less one.|
|_Size_|Optional| **Long**|The size of the dropped capital letters in number of lines high. The default is 5.|
|_Span_|Optional| **Long**|The number of letters included in the dropped capital letter. The default is 1.|
|_FontName_|Optional| **String**|The name of the font to format the dropped capital letter. The default is the current font.|
|_Bold_|Optional| **Boolean**| **True** to bold the dropped capital letter. The default is **False**.|
|_Italic_|Optional| **Boolean**| **True** to italicize the dropped capital letter. The default is **False**.|

## Example

This example formats the first three letters of the paragraphs in the specified text box.

```vb
Sub CustDropCap() 
 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.DropCap _ 
 .ApplyCustomDropCap LinesUp:=1, Size:=6, Span:=3, _ 
 FontName:="Script MT Bold", Bold:=True, Italic:=True 
 
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]