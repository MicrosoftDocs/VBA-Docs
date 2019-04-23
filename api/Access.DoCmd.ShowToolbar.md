---
title: DoCmd.ShowToolbar method (Access)
keywords: vbaac10.chm4185
f1_keywords:
- vbaac10.chm4185
ms.prod: access
api_name:
- Access.DoCmd.ShowToolbar
ms.assetid: 63663cc5-a591-c847-25c8-25777cf7806a
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.ShowToolbar method (Access)

The **ShowToolbar** method carries out the ShowToolbar action in Visual Basic.


## Syntax

_expression_.**ShowToolbar** (_ToolbarName_, _Show_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ToolbarName_|Required|**Variant**|A string expression that's the valid name of a custom toolbar you've created. If you run Visual Basic code containing the **ShowToolbar** method in a library database, Microsoft Access looks for the toolbar with this name first in the library database, and then in the current database.|
| _Show_|Optional|**[AcShowToolbar](Access.AcShowToolbar.md)**| An **AcShowToolbar** constant that specifies whether to display or hide the toolbar and in which views to display or hide it. The default value is **acToolbarYes**.|

## Remarks

You can use the **ShowToolbar** method to display or hide a custom toolbar.

If you want to show a particular toolbar on just one form or report, you can set the **OnActivate** property of the form or report to the name of a macro that contains a ShowToolbar action to show the toolbar. You can then set the **OnDeactivate** property of the form or report to the name of a macro that contains a ShowToolbar action to hide the toolbar.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
