---
title: DoCmd.NavigateTo method (Access)
keywords: vbaac10.chm5689
f1_keywords:
- vbaac10.chm5689
ms.prod: access
api_name:
- Access.DoCmd.NavigateTo
ms.assetid: 27a6e4ee-1c03-2652-3c5a-73c45f3109df
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.NavigateTo method (Access)

You can use the **NavigateTo** method to control the display of database objects in the navigation pane. 


## Syntax

_expression_.**NavigateTo** (_Category_, _Group_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Category_|Optional|**Variant**|The category by which you want the navigation pane to display objects. |
| _Group_|Optional|**Variant**|Determines which objects in the category appear in the navigation pane. If you leave this argument blank, the navigation pane will display all database objects grouped by the criteria that you specify in the _Category_ argument. Examples of valid _Group_ arguments for the various _Category_ arguments are shown in the table under **Remarks**.|

## Remarks

For example, you can change how the database objects are categorized, and you can filter the objects so that only certain ones are displayed. 

This action is similar to selecting categories and groups from the title bar of the navigation pane.

Valid _Group_ arguments depend on which _Category_ argument is used. If you enter an invalid _Group_ argument, an error message appears.

The following table contains examples of valid _Group_ arguments for each _Category_ argument.

|Category argument|Example Group arguments|
|:-----|:-----|
|Object Type|Tables, Forms, Queries, Pages, Macros, Modules|
|Tables and Views|Names of specific tables in your database|
|Modified Date|Today, Yesterday, Last Month, Older|
|Created Date|Today, Yesterday, Last Month, Older|
|Custom Category|Names of groups that you have created for the specified custom category|

> [!NOTE] 
> To navigate to the top level of a category (for example, **All Tables**, **All Access Objects**, or **All Dates**), you must leave the _Group_ argument blank. For example, when the _Category_ argument is **Object Type**, entering **All Access Objects** as a _Group_ argument results in an error.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]