---
title: Chart.PrintedCommentPages property (Excel)
keywords: vbaxl10.chm149186
f1_keywords:
- vbaxl10.chm149186
api_name:
- Excel.Chart.PrintedCommentPages
ms.assetid: 8f98f7af-4e2f-8743-b82b-c84ae83f6fdf
ms.date: 04/16/2019
ms.localizationpriority: medium
---


# Chart.PrintedCommentPages property (Excel)

Returns the number of comment pages that will be printed for the current chart. Read-only.


## Syntax

_expression_.**PrintedCommentPages**

_expression_ A variable that returns a **[Chart](Excel.Chart(object).md)** object.


## Return value

**Long**


## Remarks

Because charts and chart sheets don't support comments, the **PrintedCommentPages** property of the **Chart** object will always return zero.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]