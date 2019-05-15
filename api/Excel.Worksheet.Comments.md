---
title: Worksheet.Comments property (Excel)
keywords: vbaxl10.chm175139
f1_keywords:
- vbaxl10.chm175139
ms.prod: excel
api_name:
- Excel.Worksheet.Comments
ms.assetid: c2ad8ea7-0fa3-7cde-e3f2-49bbdb81d261
ms.date: 05/15/2019
localization_priority: Normal
---


# Worksheet.Comments property (Excel)

Returns a **[Comments](Excel.Comments.md)** collection that represents all the comments for the specified worksheet. Read-only.


## Syntax

_expression_.**Comments**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example deletes all comments added by author Jean Selva on the active sheet.

```vb
For Each c in ActiveSheet.Comments 
 If c.Author = "Jean Selva" Then c.Delete 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]