---
title: OLEObject.AutoUpdate property (Excel)
keywords: vbaxl10.chm417075
f1_keywords:
- vbaxl10.chm417075
api_name:
- Excel.OLEObject.AutoUpdate
ms.assetid: 3834c552-a282-ab75-781e-42c055346b7d
ms.date: 05/02/2019
ms.localizationpriority: medium
---


# OLEObject.AutoUpdate property (Excel)

**True** if the OLE object is updated automatically when the source changes. Valid only if the object is linked; its **OLEType** property must be **xlOLELink** (**[XlOLEType](excel.xloletype.md)** enumeration). Read-only **Boolean**.


## Syntax

_expression_.**AutoUpdate**

_expression_ A variable that represents an **[OLEObject](Excel.OLEObject.md)** object.


## Example

This example displays the status of automatic updating for all OLE objects on Sheet1.

```vb
Worksheets("Sheet1").Activate 
Range("A1").Value = "Name" 
Range("B1").Value = "Link Status" 
Range("C1").Value = "AutoUpdate Status" 
i = 2 
For Each obj In ActiveSheet.OLEObjects 
 Cells(i, 1) = obj.Name 
 If obj.OLEType = xlOLELink Then 
 Cells(i, 2) = "Linked" 
 Cells(i, 3) = obj.AutoUpdate 
 Else 
 Cells(i, 2) = "Embedded" 
 End If 
 i = i + 1 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]