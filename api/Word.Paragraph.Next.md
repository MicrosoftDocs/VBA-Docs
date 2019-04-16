---
title: Paragraph.Next method (Word)
keywords: vbawd10.chm156696900
f1_keywords:
- vbawd10.chm156696900
ms.prod: word
api_name:
- Word.Paragraph.Next
ms.assetid: 5ada0da7-a579-b728-0483-b698a09eb41c
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.Next method (Word)

Returns a  **Paragraph** object that represents the next paragraph.


## Syntax

_expression_.**Next** (_Count_)

_expression_ Required. A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Count_|Optional| **Variant**|The number of paragraphs by which you want to move ahead. The default value is one.|

## Return value

Paragraph


## Example

This example inserts a number and a tab before the first nine paragraphs in the active document.


```vb
For n = 0 To 8 
 Set myRange = ActiveDocument.Paragraphs(1).Next(Count:=n).Range 
 myRange.Collapse Direction:=wdCollapseStart 
 myRange.InsertAfter n + 1 & vbTab 
Next n
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]