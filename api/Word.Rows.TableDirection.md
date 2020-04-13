---
title: Rows.TableDirection property (Word)
keywords: vbawd10.chm155975784
f1_keywords:
- vbawd10.chm155975784
ms.prod: word
api_name:
- Word.Rows.TableDirection
ms.assetid: 02351774-13c0-ec82-c553-3b048eabb133
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.TableDirection property (Word)

Returns or sets the direction in which Microsoft Word orders cells in the specified table or row. Read/write  **[WdTableDirection](Word.WdTableDirection.md)**.


## Syntax

_expression_. `TableDirection`

_expression_ Required. A variable that represents a **[Rows](Word.Rows.md)** object.


## Remarks

If the **TableDirection** property is set to **wdTableDirectionLtr**, the selected rows are arranged with the first column in the leftmost position. If the **TableDirection** property is set to **wdTableDirectionRtl**, the selected rows are arranged with the first column in the rightmost position.


## Example

This example sets Microsoft Word to order cells in the selected row from right to left.


```vb
Selection.Rows.TableDirection = _ 
 wdTableDirectionRtl
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]