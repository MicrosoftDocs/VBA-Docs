---
title: CubeField.HasMemberProperties property (Excel)
keywords: vbaxl10.chm668086
f1_keywords:
- vbaxl10.chm668086
ms.prod: excel
api_name:
- Excel.CubeField.HasMemberProperties
ms.assetid: bd0cb9e0-95e5-47bf-3354-628bcfa604c2
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeField.HasMemberProperties property (Excel)

Returns **True** when there are member properties specified to be displayed for the cube field. Read-only **Boolean**.


## Syntax

_expression_.**HasMemberProperties**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Example

The example determines if there are member properties to be displayed for the cube field and notifies the user. The example assumes that a PivotTable exists on the active worksheet.

```vb
Sub UseHasMemberProperties() 
 
 Dim pvtTable As PivotTable 
 Dim cbeField As CubeField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set cbeField = pvtTable.CubeFields("[Country]") 
 
 ' Determine if there are member properties to be displayed. 
 If cbeField.HasMemberProperties = True Then 
 MsgBox "There are member properties to be displayed." 
 Else 
 MsgBox "There are no member properties to be displayed." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]