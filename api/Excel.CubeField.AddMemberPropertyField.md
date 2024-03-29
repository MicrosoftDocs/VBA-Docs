---
title: CubeField.AddMemberPropertyField method (Excel)
keywords: vbaxl10.chm668094
f1_keywords:
- vbaxl10.chm668094
api_name:
- Excel.CubeField.AddMemberPropertyField
ms.assetid: 721f9720-00c0-d9cf-1413-f3b0cc658595
ms.date: 04/23/2019
ms.localizationpriority: medium
---


# CubeField.AddMemberPropertyField method (Excel)

Adds a member property field to the display for the cube field.

## Syntax

_expression_.**AddMemberPropertyField** (_Property_, _PropertyOrder_, _PropertyDisplayedIn_)

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Property_|Required| **String**|The unique name of the member property. For balanced hierarchies, a unique name can be created by appending the "quoted" member property name to the unique name of the level with which the member property is associated.<br/><br/>For unbalanced hierarchies, a unique name can be created by appending the "quoted" member property name to the unique name of the hierarchy.|
| _PropertyOrder_|Optional| **Variant**|Sets the **[PropertyOrder](Excel.PivotField.PropertyOrder.md)** property value for a **CubeField** object.<br/><br/>The actual position in the collection will be immediately before the PivotTable field that currently has the same **PropertyOrder** value that is given in the argument. If no field has the given **PropertyOrder** value, the range of acceptable values is 1 to the number of member properties already showing for the hierarchy plus one.<br/><br/>This argument is one-based. If omitted, the property goes to the end of the list.|
| _PropertyDisplayedIn_|Optional| **[XlPropertyDisplayedIn](Excel.XlPropertyDisplayedIn.md)**|Specifies where to display the property. If this argument is omitted, the member property field will be added to the PivotTable only.|

## Remarks

The property field specified will not be viewable if the PivotTable view has no fields.

To delete member properties, use the **Delete** method to delete the **PivotField** object from the **PivotFields** collection.


## Example

In this example, Microsoft Excel adds a member property field titled Description to the PivotTable report view. This example assumes that a PivotTable exists on the active worksheet, and that Country, Area, and Description are items in the report.

```vb
Sub UseAddMemberPropertyField() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 With pvtTable 
 .ManualUpdate = True 
 .CubeFields("[Country]").LayoutForm = xlOutline 
 .CubeFields("[Country]").AddMemberPropertyField _ 
 Property:="[Country].[Area].[Description]" 
 .ManualUpdate = False 
 End With 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]