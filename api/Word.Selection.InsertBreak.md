---
title: Selection.InsertBreak method (Word)
keywords: vbawd10.chm158662778
f1_keywords:
- vbawd10.chm158662778
ms.prod: word
api_name:
- Word.Selection.InsertBreak
ms.assetid: 2c9d8cb8-1cc1-3d69-1e26-3a6878c0b1da
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.InsertBreak method (Word)

Inserts a page, column, or section break.


## Syntax

_expression_. `InsertBreak`( `_Type_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[WdBreakType](Word.WdBreakType.md)**|the type of break to insert. The default value is **wdPageBreak**. Some of the **WdBreakType** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|

## Remarks

When you insert a page or column break, the break replaces the selection. If you don't want to replace the selection, use the **[Collapse](Word.Selection.Collapse.md)** method before using the **InsertBreak** method.


> [!NOTE] 
> When you insert a section break, the break is inserted immediately preceding the selection.


## Example

This example inserts a continuous section break immediately preceding the selection.


```vb
Selection.InsertBreak Type:=wdSectionBreakContinuous
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
