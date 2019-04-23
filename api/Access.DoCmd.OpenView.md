---
title: DoCmd.OpenView method (Access)
keywords: vbaac10.chm4649
f1_keywords:
- vbaac10.chm4649
ms.prod: access
api_name:
- Access.DoCmd.OpenView
ms.assetid: 8d2970dd-9a06-f917-04da-850b085126dd
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.OpenView method (Access)

The **OpenView** method carries out the OpenView action in Visual Basic.


## Syntax

_expression_.**OpenView** (_ViewName_, _View_, _DataMode_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ViewName_|Required|**Variant**|A string expression that's the valid name of a view in the current database. If you execute Visual Basic code containing the **OpenView** method in a library database, Microsoft Access looks for the view with this name first in the library database, and then in the current database.|
| _View_|Optional|**[AcView](Access.AcView.md)**|An **AcView** constant that specifies the view in which the view will open. The default value is **acViewNormal**.|
| _DataMode_|Optional|**[AcOpenDataMode](Access.AcOpenDataMode.md)**|An **AcOpenDataMode** constant that specifies the data entry mode for the view. The default value is **acEdit**.|

## Remarks

In an Access project, you can use the **OpenView** method to open a view in Datasheet view, view Design view, or Print Preview. This action runs the named view when opened in Datasheet view. You can select data entry for the view and restrict the records that the view displays.

If you don't want to display the system messages that normally appear when a view is run (indicating it's a view and showing how many records will be affected), you can use the **SetWarnings** method to suppress the display of these messages.


## Example

The following example opens the Employees view.

```vb
DoCmd.OpenView "Employees"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]