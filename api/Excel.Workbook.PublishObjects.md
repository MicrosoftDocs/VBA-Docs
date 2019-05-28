---
title: Workbook.PublishObjects property (Excel)
keywords: vbaxl10.chm199187
f1_keywords:
- vbaxl10.chm199187
ms.prod: excel
api_name:
- Excel.Workbook.PublishObjects
ms.assetid: b6418f80-5154-6e3f-7313-222e6438c0e1
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.PublishObjects property (Excel)

Returns the **[PublishObjects](Excel.PublishObjects.md)** collection. Read-only.


## Syntax

_expression_.**PublishObjects**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example publishes all static **PublishObject** objects in the active workbook to the webpage.

```vb
Set objPObjs = ActiveWorkbook.PublishObjects 
For Each objPO in objPObjs 
 If objPO.HtmlType = xlHTMLStatic Then 
 objPO.Publish 
 End If 
Next objPO
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]