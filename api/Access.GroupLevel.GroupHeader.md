---
title: GroupLevel.GroupHeader property (Access)
keywords: vbaac10.chm12241
f1_keywords:
- vbaac10.chm12241
ms.prod: access
api_name:
- Access.GroupLevel.GroupHeader
ms.assetid: 336e45dc-595e-a3fd-9d6b-9e1153654602
ms.date: 03/20/2019
localization_priority: Normal
---


# GroupLevel.GroupHeader property (Access)

You can use the **GroupHeader** property to create a group header for a selected field or expression in a report. Read/write **Boolean**.


## Syntax

_expression_.**GroupHeader**

_expression_ A variable that represents a **[GroupLevel](Access.GroupLevel.md)** object.


## Remarks

You can use group headers and footers to label or summarize data in a group of records. For example, if you set the **GroupHeader** property to Yes for the **Categories** field, each group of products will begin with its category name.

> [!NOTE] 
> You can't set or refer to these properties directly in Visual Basic. To create a group header or footer for a field or expression in Visual Basic, use the **[CreateGroupLevel](Access.Application.CreateGroupLevel.md)** method.

To set the grouping properties—**[GroupOn](Access.GroupLevel.GroupOn.md)**, **[GroupInterval](Access.GroupLevel.GroupInterval.md)**, and **[KeepTogether](Access.GroupLevel.KeepTogether.md)**—to other than their default values, you must first set the **GroupHeader** or **GroupFooter** property or both to Yes for the selected field or expression.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]