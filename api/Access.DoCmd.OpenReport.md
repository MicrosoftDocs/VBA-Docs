---
title: DoCmd.OpenReport method (Access)
keywords: vbaac10.chm4163
f1_keywords:
- vbaac10.chm4163
ms.prod: access
api_name:
- Access.DoCmd.OpenReport
ms.assetid: 3c08755a-5116-f085-d498-725dc12e62f1
ms.date: 03/07/2019
localization_priority: Priority
---


# DoCmd.OpenReport method (Access)

The **OpenReport** method carries out the OpenReport action in Visual Basic.


## Syntax

_expression_.**OpenReport** (_ReportName_, _View_, _FilterName_, _WhereCondition_, _WindowMode_, _OpenArgs_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ReportName_|Required|**Variant**|A string expression that's the valid name of a report in the current database. If you execute Visual Basic code containing the **OpenReport** method in a library database, Microsoft Access looks for the report with this name first in the library database, and then in the current database.|
| _View_|Optional|**[AcView](Access.AcView.md)**|An **AcView** constant that specifies the view in which the report will open. The default value is **acViewNormal**.|
| _FilterName_|Optional|**Variant**|A string expression that's the valid name of a query in the current database.|
| _WhereCondition_|Optional|**Variant**|A string expression that's a valid SQL WHERE clause without the word WHERE.|
| _WindowMode_|Optional|**[AcWindowMode](Access.AcWindowMode.md)**|An **AcWindowMode** constant that specifies the mode in which the form opens. The default value is **acWindowNormal**.|
| _OpenArgs_|Optional|**Variant**|Sets the **OpenArgs** property.|

## Remarks

You can use the **OpenReport** method to open a report in Design view or Print Preview, or to print the report immediately. You can also restrict the records that are printed in the report.

The maximum length of the _WhereCondition_ argument is 32,768 characters (unlike the _WhereCondition_ action argument in the Macro window, whose maximum length is 256 characters).


## Example

The following example prints Sales Report while using the existing query Report Filter.

```vb
DoCmd.OpenReport "Sales Report", acViewNormal, "Report Filter"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
