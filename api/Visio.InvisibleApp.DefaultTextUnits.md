---
title: InvisibleApp.DefaultTextUnits property (Visio)
keywords: vis_sdr.chm17551035
f1_keywords:
- vis_sdr.chm17551035
ms.prod: visio
api_name:
- Visio.InvisibleApp.DefaultTextUnits
ms.assetid: a9bd8464-b39c-622c-6446-bc652e42766c
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.DefaultTextUnits property (Visio)

Determines the default unit of measure for quantities that represent text metrics. Read/write.


## Syntax

_expression_.**DefaultTextUnits**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

Variant


## Remarks

The **DefaultTextUnits** property corresponds to the value shown in the **Text** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box (**File** tab > **Options**).

The return value contains one of the values of **[VisUnitCodes](Visio.visunitcodes.md)**, which are declared in the Microsoft Visio type library.

You can specify the value of **DefaultTextUnits** as an integer (a member of **VisUnitCodes**) or a string value such as "pt". If the string is invalid or the unit code is inappropriate (non-textual), an error is generated.

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About units of measure](../visio/Concepts/about-units-of-measure-visio.md).

Cell formulas that contain a specific unit of measure are displayed in those units regardless of the default text units setting. Many cell formulas, however, use implicit unit syntax and are displayed in default units.

A program can create a cell whose formula is displayed in default units by setting the cell's **Formula** property to a string in implicit unit syntax. For example, the formula `"=8[pt,T]"` is displayed as `"8 pt"` if the **DefaultTextUnits** property is **visPoints** and as `"0.6272"` if the **DefaultTextUnits** property is **visCiceros**.

Alternatively, a program can use the following statement to set the cell's result to default text units. 

```vb
vsoCell.Result(visTextUnits) = 12
```

In this case, the text is 12 points if the **DefaultTextUnits** property is **visPoints** and 12 ciceros if the **DefaultTextUnits** property is **visCiceros**.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]