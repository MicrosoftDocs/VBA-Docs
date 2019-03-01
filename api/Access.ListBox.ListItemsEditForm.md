---
title: ListBox.ListItemsEditForm property (Access)
keywords: vbaac10.chm11336
f1_keywords:
- vbaac10.chm11336
ms.prod: access
api_name:
- Access.ListBox.ListItemsEditForm
ms.assetid: f744fc52-4c50-f740-7a2f-eeccb12de7c9
ms.date: 03/02/2019
localization_priority: Normal
---


# ListBox.ListItemsEditForm property (Access)

Gets or sets the name of the form that is displayed when the user chooses **Edit List Items**. Read/write **String**.


## Syntax

_expression_.**ListItemsEditForm**

_expression_ A variable that represents a **[ListBox](Access.ListBox.md)** object.


## Remarks

The **AllowValueEditLists** property determines whether the **Edit List Items** command is available when the user right-clicks a list box that's bound to a **Lookup** field.

If the **Lookup** field is bound to a table or query, the form specified by the **ListItemsEditForm** property is displayed when the user chooses **Edit List Items**. The user can use the form to add, delete, or edit the items to be displayed in the list box.

The **ListItemsEditForm** property is not available for list boxes on a report.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]