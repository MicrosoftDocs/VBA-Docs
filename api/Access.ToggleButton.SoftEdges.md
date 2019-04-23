---
title: ToggleButton.SoftEdges property (Access)
keywords: vbaac10.chm14639
f1_keywords:
- vbaac10.chm14639
ms.prod: access
api_name:
- Access.ToggleButton.SoftEdges
ms.assetid: 23c63821-966c-4d9f-7304-5b6e31b85675
ms.date: 03/05/2019
localization_priority: Normal
---


# ToggleButton.SoftEdges property (Access)

Gets or sets the soft edges effect applied to the specified object. Read/write **Long**.


## Syntax

_expression_.**Soft Edges**

_expression_ A variable that represents a **[ToggleButton](Access.ToggleButton.md)** object.


## Remarks

The **SoftEdges** property uses one of the values listed in the following table.

|Value|Effect|
|:----|:-----|
|0 (Default)|No Soft Edges|
|1|1 Point|
|2|2.5 Points|
|3|5 Points|
|4|10 Points|
|5|25 Points|
|6|50 Points|

To see the available soft edges effects and apply soft edges through the user interface, first open the form or report in Layout view or Design view by right-clicking the form or report in the navigation pane, and then choosing the view that you want. 

Next, choose the object to which you want to apply a soft edges effect. On the **Format** tab, in the **Control Formatting** group, choose **Shape Effects** > **Soft Edges**, and then choose a soft edges effect. Notice that the soft edges effects are indexed from top to bottom.

This property is not surfaced in the property sheet.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]