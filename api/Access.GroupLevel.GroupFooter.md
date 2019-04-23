---
title: GroupLevel.GroupFooter property (Access)
keywords: vbaac10.chm12242
f1_keywords:
- vbaac10.chm12242
ms.prod: access
api_name:
- Access.GroupLevel.GroupFooter
ms.assetid: c10d30b2-da18-cd6f-8b00-e964cd4751d6
ms.date: 03/20/2019
localization_priority: Normal
---


# GroupLevel.GroupFooter property (Access)

You can use the **GroupFooter** property to create a group footer for a selected field or expression in a report. Read/write **Boolean**.


## Syntax

_expression_.**GroupFooter**

_expression_ A variable that represents a **[GroupLevel](Access.GroupLevel.md)** object.


## Remarks

You can use group headers and footers to label or summarize data in a group of records. For example, if you set the **GroupHeader** property to Yes for the **Categories** field, each group of products will begin with its category name.

> [!NOTE] 
> You can't set or refer to these properties directly in Visual Basic. To create a group header or footer for a field or expression in Visual Basic, use the **[CreateGroupLevel](Access.Application.CreateGroupLevel.md)** method.

To set the grouping properties—**[GroupOn](Access.GroupLevel.GroupOn.md)**, **[GroupInterval](Access.GroupLevel.GroupInterval.md)**, and **[KeepTogether](Access.GroupLevel.KeepTogether.md)**—to other than their default values, you must first set the **GroupHeader** or **GroupFooter** property or both to Yes for the selected field or expression.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]