---
title: CommentThreaded.Text method (Excel)
keywords:
f1_keywords:
-
ms.prod: excel
api_name:
- Excel.CommentThreaded.Text
ms.assetid:
ms.date: 05/08/2019
localization_priority: Normal
---


# CommentThreaded.Text method (Excel)

Sets CommentThreaded text.


## Syntax

_expression_.**Text** (_Text_, _Start_, _Overwrite_)

_expression_ A variable that represents a **[CommentThreaded](Excel.CommentThreaded.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Optional| **Variant**|The text to be added.|
| _Start_|Optional| **Variant**|The character number where the added text will be placed. If this argument is omitted, any existing text in the comment is deleted.|
| _Overwrite_|Optional| **Variant**| **True** to overwrite the existing text. The default value is **False** (text is inserted).|

## Return value

String




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
