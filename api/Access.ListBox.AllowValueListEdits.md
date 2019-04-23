---
title: ListBox.AllowValueListEdits property (Access)
keywords: vbaac10.chm11335
f1_keywords:
- vbaac10.chm11335
ms.prod: access
api_name:
- Access.ListBox.AllowValueListEdits
ms.assetid: cab2ec6f-affb-5111-af5e-6f3638189dff
ms.date: 02/28/2019
localization_priority: Normal
---


# ListBox.AllowValueListEdits property (Access)

Gets or sets whether the **Edit List Items** command is available when the user right-clicks a list box. Read/write **Boolean**.


## Syntax

_expression_.**AllowValueEditLists**

_expression_ A variable that represents a **[ListBox](Access.ListBox.md)** object.


## Remarks

The **AllowValueEditLists** property determines whether the **Edit List Items** command is available when the user right-clicks a list box that's bound to a **Lookup** field.

If the **Lookup** field is bound to a list of values, the **Edit List Items** dialog box is displayed when the user chooses **Edit List Items**. The user can then add, delete, or edit the items to be displayed in the list box.

If the **Lookup** field is bound to a table or query, the form specified by the **ListItemsEditForm** property is displayed when the user chooses **Edit List Items**. The user can use the form to add, delete, or edit the items to be displayed in the list box.

The **AllowValueEditLists** property is not available for list boxes on a report.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

