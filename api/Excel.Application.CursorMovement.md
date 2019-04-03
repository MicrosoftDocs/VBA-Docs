---
title: Application.CursorMovement property (Excel)
keywords: vbaxl10.chm133237
f1_keywords:
- vbaxl10.chm133237
ms.prod: excel
api_name:
- Excel.Application.CursorMovement
ms.assetid: 4be5a3fd-7a68-1190-5888-239497d53cb1
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.CursorMovement property (Excel)

Returns or sets a value that indicates whether a visual cursor or a logical cursor is used. Can be one of the following constants: **xlVisualCursor** or **xlLogicalCursor**. Read/write **Long**.


## Syntax

_expression_.**CursorMovement**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

These constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example sets Microsoft Excel to use the visual cursor.

```vb
Application.CursorMovement = xlVisualCursor
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]