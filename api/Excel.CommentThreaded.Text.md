---
title: CommentThreaded.Text method (Excel)
keywords: vbaxl10.chm1010075
f1_keywords:
- vbaxl10.chm1010075
ms.prod: excel
api_name:
- Excel.CommentThreaded.Text
ms.date: 06/27/2019
localization_priority: Normal
---


# CommentThreaded.Text method (Excel)

Sets threaded comment text.


## Syntax

_expression_.**Text** (_Text_, _Start_, _Overwrite_)

_expression_ A variable that represents a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Optional| **Variant**|The text to be added.|
| _Start_|Optional| **Variant**|The character number where the added text will be placed. If the _Overwrite_ parameter is **True** or blank, and if this argument is omitted, any existing text in the threaded comment is deleted.|
| _Overwrite_|Optional| **Variant**| **False** to insert the text. The default value is **True** (text is overwritten).|


## Return value

**String**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
