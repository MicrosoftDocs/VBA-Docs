---
title: ComboBox.RemoveItem Method (Outlook Forms Script)
keywords: olfm10.chm2000370
f1_keywords:
- olfm10.chm2000370
ms.prod: outlook
ms.assetid: abbc1126-4983-a583-0fd4-b76418d5c2cb
ms.date: 06/08/2017
localization_priority: Normal
---


# ComboBox.RemoveItem Method (Outlook Forms Script)

Removes a row from the list in a  **[ComboBox](Outlook.combobox.md)**.


## Syntax

_expression_.**RemoveItem**(**_pvargIndex_**)

_expression_ A variable that represents a  **ComboBox** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|pvargIndex|Required| **Variant**|Specifies the row to delete. The number of the first row is 0; the number of the second row is 1, and so on.|

## Return value

A  **Boolean** that returns **True** if the method succeeds, **False** otherwise.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]