---
title: ServerViewableItems.Delete method (Excel)
keywords: vbaxl10.chm833075
f1_keywords:
- vbaxl10.chm833075
ms.prod: excel
api_name:
- Excel.ServerViewableItems.Delete
ms.assetid: e6b53271-8a37-4bf3-fea2-46d02550391b
ms.date: 05/14/2019
localization_priority: Normal
---


# ServerViewableItems.Delete method (Excel)

Deletes a reference to an object in the **ServerViewableItems** collection in the workbook.


## Syntax

_expression_.**Delete** (_Index_)

_expression_ A variable that represents a **[ServerViewableItems](Excel.ServerViewableItems.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index of the object that you want to delete.|

## Remarks

If you do not want a particular object to be viewable in Excel Services, use this method to remove that object.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]