---
title: PivotTable.TableStyle2 property (Excel)
keywords: vbaxl10.chm235171
f1_keywords:
- vbaxl10.chm235171
ms.prod: excel
api_name:
- Excel.PivotTable.TableStyle2
ms.assetid: d2d79fc6-2ead-91a9-f304-92248584f4b2
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.TableStyle2 property (Excel)

The **TableStyle2** property specifies the PivotTable style currently applied to the PivotTable. Read/write.


## Syntax

_expression_.**TableStyle2**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

The property is called **TableStyle2** because there is an existing property named **[TableStyle](excel.listobject.tablestyle.md)**.


## Example

```vb
Sub ApplyingStyle() 
 
 ActiveSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight17" 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]