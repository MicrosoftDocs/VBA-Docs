---
title: ColumnFormat.Width property (Outlook)
keywords: vbaol11.chm2730
f1_keywords:
- vbaol11.chm2730
ms.prod: outlook
api_name:
- Outlook.ColumnFormat.Width
ms.assetid: d0dd6c11-bce4-3785-7686-7863466d2380
ms.date: 06/08/2017
localization_priority: Normal
---


# ColumnFormat.Width property (Outlook)

Returns or sets a  **Long** value indicating the approximate width (in characters) of the column. Read/write.


## Syntax

_expression_.**Width**

_expression_ A variable that represents a [ColumnFormat](Outlook.ColumnFormat.md) object.


## Remarks

This property can be set to a value between 2 and 1024. If this property is set to a value less than 2, the property is set to 2. If this property is set to a value greater than 1024, the property is set to 1024.

If the value of this property for every column in a view is less than the total width of the view, then the  **Width** property of the **[ColumnFormat](Outlook.ColumnFormat.md)** object for the last **[ViewField](Outlook.ViewField.md)** in the view is increased to match the total width of the view.


## See also


[ColumnFormat Object](Outlook.ColumnFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]