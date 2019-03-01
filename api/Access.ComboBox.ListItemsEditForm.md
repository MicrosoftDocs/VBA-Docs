---
title: ComboBox.ListItemsEditForm property (Access)
keywords: vbaac10.chm11518
f1_keywords:
- vbaac10.chm11518
ms.prod: access
api_name:
- Access.ComboBox.ListItemsEditForm
ms.assetid: 5db884d4-4d9f-23b5-9e3a-f6de953a4800
ms.date: 03/02/2019
localization_priority: Normal
---


# ComboBox.ListItemsEditForm property (Access)

Gets or sets the name of the form that is displayed when the user chooses **Edit List Items**. Read/write **String**.


## Syntax

_expression_.**ListItemsEditForm**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

The **AllowValueEditLists** property determines whether the **Edit List Items** command is available when the user right-clicks a combo box that's bound to a **Lookup** field.

If the **Lookup** field is bound to a table or query, the form specified by the **ListItemsEditForm** property is displayed when the user chooses **Edit List Items**. The user can use the form to add, delete, or edit the items to be displayed in the combo box.

The **ListItemsEditForm** property is not available for combo boxes on a report.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]