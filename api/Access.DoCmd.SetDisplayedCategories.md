---
title: DoCmd.SetDisplayedCategories method (Access)
keywords: vbaac10.chm5851
f1_keywords:
- vbaac10.chm5851
ms.prod: access
api_name:
- Access.DoCmd.SetDisplayedCategories
ms.assetid: ae2290c3-43ff-c19d-63f8-41427aacd9ce
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.SetDisplayedCategories method (Access)

Specifies which categories are displayed under **Navigate to Category** in the title bar of the navigation pane. 


## Syntax

_expression_.**SetDisplayedCategories** (_Show_, _Category_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Show_|Required|**Variant**|Set to Yes to show the category or categories. Set to No to hide them.|
| _Category_|Optional|**Variant**|The name of the category that you want to show or hide. Leave blank to show or hide all categories.|

## Remarks

For example, if you want to prevent users from switching the navigation pane so that it displays objects sorted by **Created Date**, you can use this method to hide that option in the title bar's drop-down list.

The caption in the title bar of the navigation pane indicates which filter, if any, is currently active. Click anywhere in the bar to display the drop-down list. The items that this method controls are listed under **Navigate to Category**.

This method only enables or disables the display of the specified category or categories; it does not perform any switching of the navigation pane display. For example, if you are displaying objects sorted by **Creation Date** and you use the **SetDisplayedCategories** method to disable the **Creation Date** option, Access does not switch the navigation pane to another category.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]