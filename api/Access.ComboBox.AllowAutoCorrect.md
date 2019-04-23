---
title: ComboBox.AllowAutoCorrect property (Access)
keywords: vbaac10.chm11397
f1_keywords:
- vbaac10.chm11397
ms.prod: access
api_name:
- Access.ComboBox.AllowAutoCorrect
ms.assetid: ebf48367-20fb-14be-7082-a2d9de923c51
ms.date: 02/28/2019
localization_priority: Normal
---


# ComboBox.AllowAutoCorrect property (Access)

You can use the **AllowAutoCorrect** property to specify whether the specified control will automatically correct entries made by the user. Read/write **Boolean**.


## Syntax

_expression_.**AllowAutoCorrect**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

The **AllowAutoCorrect** property will correct spelling errors in filters for combo boxes with **LimitToList** set to **True** (for example, INitial CAps as set by the Autocorrect settings of the database), and should be used with caution.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]