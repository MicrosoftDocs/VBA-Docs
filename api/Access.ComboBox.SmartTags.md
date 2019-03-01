---
title: ComboBox.SmartTags property (Access)
keywords: vbaac10.chm11478
f1_keywords:
- vbaac10.chm11478
ms.prod: access
api_name:
- Access.ComboBox.SmartTags
ms.assetid: b86a8460-48c6-92ad-602b-1d736bb2c38c
ms.date: 03/02/2019
localization_priority: Normal
---


# ComboBox.SmartTags property (Access)

Returns a **[SmartTags](Access.SmartTags.md)** collection that represents the collection of smart tags that have been added to a control. 


## Syntax

_expression_.**SmartTags**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

Unlike the **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0)` returns the first smart tag for the specified control.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]