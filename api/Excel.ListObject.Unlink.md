---
title: ListObject.Unlink method (Excel)
keywords: vbaxl10.chm734076
f1_keywords:
- vbaxl10.chm734076
ms.prod: excel
api_name:
- Excel.ListObject.Unlink
ms.assetid: 37e70576-e577-cfbb-f5ee-63ba830e174e
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.Unlink method (Excel)

Removes the link to a Microsoft SharePoint Foundation site from a list. Returns **Nothing**.


## Syntax

_expression_.**Unlink**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Remarks

After this method is called and the list is unlinked, it cannot be reversed.


## Example

The following example unlinks a list from a SharePoint site.

```vb
Sub UnlinkList() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 
 objListObj.Unlink 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]