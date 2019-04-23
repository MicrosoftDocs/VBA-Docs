---
title: Comment.Shape property (Excel)
keywords: vbaxl10.chm516074
f1_keywords:
- vbaxl10.chm516074
ms.prod: excel
api_name:
- Excel.Comment.Shape
ms.assetid: f3e5f713-69b3-9cd8-81fa-9c677ed26869
ms.date: 04/23/2019
localization_priority: Normal
---


# Comment.Shape property (Excel)

Returns a **[Shape](Excel.Shape.md)** object that represents the shape attached to the specified comment.


## Syntax

_expression_.**Shape**

_expression_ A variable that represents a **[Comment](Excel.Comment.md)** object.


## Example

This example selects comment two on the active sheet.

> [!NOTE] 
> Ensure that the comments are not hidden. On the **Review** tab, choose **Comments** > **Show All Comments**.

```vb
ActiveSheet.Comments(2).Shape.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]