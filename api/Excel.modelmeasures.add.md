---
title: ModelMeasures.Add method (Excel)
keywords: vbaxl10.chm980077
f1_keywords:
- vbaxl10.chm980077
ms.assetid: abc0f260-abdb-2f60-928f-b325fbb976f3
ms.date: 05/01/2019
ms.prod: excel
localization_priority: Normal
---


# ModelMeasures.Add method (Excel)

Adds a model measure to the model.


## Syntax

_expression_.**Add** (_MeasureName_, _AssociatedTable_, _Formula_, _FormatInformation_, _Description_)

_expression_ A variable that represents a **[ModelMeasures](Excel.modelmeasures.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MeasureName_|Required|**String**|The name of the model measure.|
| _AssociatedTable_|Required|MODELTABLE|The model table associated with the model measure. This is the table that contains the model measure, as seen in the **Field List** task pane.|
| _Formula_|Required|**String**|The Data Analysis Expressions (DAX) formula, inserted as a string.|
| _FormatInformation_|Required|**Variant**|The formatting of the model measure. See Remarks. |
| _Description_|Optional|**Variant**|The description associated with the model measure.|

## Return value

**[ModelMeasure](Excel.modelmeasure.md)**


## Remarks

The formatting of the model measure can be of type:

- **[ModelFormatBoolean](Excel.modelformatboolean.md)**
- **[ModelFormatCurrency](Excel.modelformatcurrency.md)**
- **[ModelFormatDate](Excel.modelformatdate.md)**
- **[ModelFormatDecimalNumber](Excel.modelformatdecimalnumber.md)**
- **[ModelFormatGeneral](Excel.modelformatgeneral.md)**
- **[ModelFormatPercentageNumber](Excel.modelformatpercentagenumber.md)**
- **[ModelFormatScientificNumber](Excel.modelformatscientificnumber.md)**
- **[ModelFormatWholeNumber](Excel.modelformatwholenumber.md)**



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]