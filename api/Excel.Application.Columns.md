---
title: Application.Columns property (Excel)
keywords: vbaxl10.chm183087
f1_keywords:
- vbaxl10.chm183087
ms.prod: excel
api_name:
- Excel.Application.Columns
ms.assetid: 242d9112-9352-c3a6-e23e-59aec3d8f68f
ms.date: 06/08/2017
localization_priority: Priority
---


# Application.Columns property (Excel)

Returns a  **[Range](Excel.Range(object).md)** object that represents all the columns on the active worksheet. If the active document isn't a worksheet, the **Columns** property fails.


## Syntax

_expression_. `Columns`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Remarks

Using this property without an object qualifier is equivalent to using  `ActiveSheet.Columns`.

When applied to a  **Range** object that's a multiple-area selection, this property returns columns from only the first area of the range. For example, if the **Range** object has two areas — A1:B2 and C3:D4 — `Selection.Columns.Count` returns 2, not 4. To use this property on a range that may contain a multiple-area selection, test `Areas.Count` to determine whether the range contains more than one area. If it does, loop over each area in the range.


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]