---
title: ComboBox.AllowAutoCorrect property (Access)
keywords: vbaac10.chm11397
f1_keywords:
- vbaac10.chm11397
ms.prod: access
api_name:
- Access.ComboBox.AllowAutoCorrect
ms.assetid: ebf48367-20fb-14be-7082-a2d9de923c51
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.AllowAutoCorrect property (Access)

You can use the  **AllowAutoCorrect** property to specify whetherthe specified control will automatically correct entries made by the user. Read/write **Boolean**.


## Syntax

_expression_. `AllowAutoCorrect`

_expression_ A variable that represents a [ComboBox](Access.ComboBox.md) object.


## Remarks

The **AllowAutoCorrect** property will correct spelling errors in filters for ComboBoxes with **LimitToList** set to True , e.g INitial CAps as set by the Autocorrect Settings of the database and should be used with caution.


## See also


[ComboBox Object](Access.ComboBox.md)

