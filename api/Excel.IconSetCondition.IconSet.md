---
title: IconSetCondition.IconSet property (Excel)
keywords: vbaxl10.chm812087
f1_keywords:
- vbaxl10.chm812087
ms.prod: excel
api_name:
- Excel.IconSetCondition.IconSet
ms.assetid: 8e0529d5-1c15-744e-2391-7229bcbcd043
ms.date: 04/27/2019
localization_priority: Normal
---


# IconSetCondition.IconSet property (Excel)

Returns or sets an **[IconSets](Excel.IconSets.md)** collection, which specifies the icon set used in the conditional format.


## Syntax

_expression_.**IconSet**

_expression_ A variable that represents an **[IconSetCondition](Excel.IconSetCondition.md)** object.


## Remarks

You can assign the icon set by using the **[IconSets](Excel.Workbook.IconSets.md)** property of the **Workbook** object. 

For example, `Selection.FormatConditions(1).IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)` will apply the three-traffic-light icon set to the conditional format.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]