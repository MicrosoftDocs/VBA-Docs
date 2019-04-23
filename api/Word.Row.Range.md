---
title: Row.Range property (Word)
keywords: vbawd10.chm156237824
f1_keywords:
- vbawd10.chm156237824
ms.prod: word
api_name:
- Word.Row.Range
ms.assetid: 1ca11d5e-9f2d-fd9f-c3a4-100e99a3f955
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.Range property (Word)

Returns a  **Range** object that represents the portion of a document that is contained within the specified table row.


## Syntax

_expression_.**Range**

_expression_ Required. A variable that represents a '[Row](Word.Row.md)' object.


## Example

This example copies the first row in table one.


```vb
If ActiveDocument.Tables.Count >= 1 Then _ 
 ActiveDocument.Tables(1).Rows(1).Range.Copy
```


## See also


[Row Object](Word.Row.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]