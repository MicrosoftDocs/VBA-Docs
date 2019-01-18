---
title: Style.Type property (Word)
keywords: vbawd10.chm153878531
f1_keywords:
- vbawd10.chm153878531
ms.prod: word
api_name:
- Word.Style.Type
ms.assetid: 2f065484-a7ec-0833-340f-93cbe954e4ba
ms.date: 06/08/2017
localization_priority: Normal
---


# Style.Type property (Word)

Returns the style type. Read-only  **[WdStyleType](Word.WdStyleType.md)**.


## Syntax

 _expression_. `Type`

 _expression_ Required. A variable that represents a '[Style](Word.Style.md)' object.


## Example

This example displays a message that indicates the style type of the style named "SubTitle" in the active document.


```vb
If ActiveDocument.Styles("SubTitle").Type = _ 
 wdStyleTypeParagraph Then 
 MsgBox "Paragraph style" 
ElseIf ActiveDocument.Styles("SubTitle").Type = _ 
 wdStyleTypeCharacter Then 
 MsgBox "Character style" 
End If
```


## See also


[Style Object](Word.Style.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]