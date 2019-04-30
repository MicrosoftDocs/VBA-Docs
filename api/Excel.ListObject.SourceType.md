---
title: ListObject.SourceType property (Excel)
keywords: vbaxl10.chm734093
f1_keywords:
- vbaxl10.chm734093
ms.prod: excel
api_name:
- Excel.ListObject.SourceType
ms.assetid: 17c41741-1bca-0c07-d113-fd68ba7add75
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.SourceType property (Excel)

Returns an **[XlListObjectSourceType](Excel.XlListObjectSourceType.md)** value that represents the current source of the list.


## Syntax

_expression_.**SourceType**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Example

The following sample code returns an **XlListObjectSourceType** constant indicating the source of the default list on Sheet1 of the active workbook.

```vb
Sub Test () 
 Dim wrksht As Worksheet 
 Dim oListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListObj = wrksht.ListObjects(1) 
 
 Debug.Print oListObj.SourceType 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]