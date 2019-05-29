---
title: Font.DisableCharacterSpaceGrid property (Word)
keywords: vbawd10.chm156369051
f1_keywords:
- vbawd10.chm156369051
ms.prod: word
api_name:
- Word.Font.DisableCharacterSpaceGrid
ms.assetid: b608a477-03a2-c1e0-eaa0-841a12665865
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.DisableCharacterSpaceGrid property (Word)

 **True** if Microsoft Word ignores the number of characters per line for the corresponding **Font** object. Read/write **Boolean**.


## Syntax

_expression_. `DisableCharacterSpaceGrid`

_expression_ A variable that represents a **[Font](Word.Font.md)** object.


## Remarks

This property returns  **wdUndefined** if the **DisableCharacterSpaceGrid** property is set to **True** for only some of the specified text.


## Example

This example signals Microsoft Word to ignore the number of characters per line for the selected text.


```vb
With Selection.Font 
 .DisableCharacterSpaceGrid = True 
End With
```


## See also


[Font Object](Word.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]