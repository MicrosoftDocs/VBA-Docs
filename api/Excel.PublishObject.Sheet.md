---
title: PublishObject.Sheet property (Excel)
keywords: vbaxl10.chm652076
f1_keywords:
- vbaxl10.chm652076
ms.prod: excel
api_name:
- Excel.PublishObject.Sheet
ms.assetid: 37aedf9e-01e1-0790-d141-6d2490e3eab2
ms.date: 05/09/2019
localization_priority: Normal
---


# PublishObject.Sheet property (Excel)

Returns the sheet name for the specified **PublishObject** object. Read-only **String**.


## Syntax

_expression_.**Sheet**

_expression_ A variable that represents a **[PublishObject](Excel.PublishObject.md)** object.



## Example

This example determines the name of the worksheet that contains the first **PublishObject** object that is saved as static HTML on the webpage. The example then sets the **Boolean** variable `blnSheetFound` to **True**. If no items in the document have been saved as static HTML, `blnSheetFound` is **False**.

```vb
blnSheetFound = False 
For Each objPO In Workbooks(1).PublishObjects 
 If objPO.HtmlType = xlHTMLStatic Then 
 strFirstPO = objPO.Sheet 
 blnSheetFound = True 
 Exit For 
 End If 
Next objPO 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]