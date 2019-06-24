---
title: InvisibleApp.DefaultAngleUnits property (Visio)
keywords: vis_sdr.chm17551050
f1_keywords:
- vis_sdr.chm17551050
ms.prod: visio
api_name:
- Visio.InvisibleApp.DefaultAngleUnits
ms.assetid: 5c7f775c-9e2b-10e0-cbc0-2ac0b922ed1a
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.DefaultAngleUnits property (Visio)

Determines the default unit of measure for quantities that represent angles. Read/write.


## Syntax

_expression_.**DefaultAngleUnits**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

Variant


## Remarks

The **DefaultAngleUnits** property corresponds to the value shown in the **Angle** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box (**File** tab > **Options**).

The return value contains one of the values of **[VisUnitCodes](Visio.visunitcodes.md)**, which are declared in the Microsoft Visio type library.

You can specify the value of the **DefaultAngleUnits** property as an integer (a member of **VisUnitCodes**) or a string value such as "degrees". If the string is invalid or the unit code is inappropriate (non-angular), an error is generated.

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About units of measure](../visio/Concepts/about-units-of-measure-visio.md).

Cell formulas that contain a specific unit of measure are displayed in those units regardless of the default angle units setting. Many cell formulas, however, use implicit unit syntax and are displayed in default units.

A program can create a cell whose formula is displayed in default units by setting the cell's **Formula** property to a string in implicit unit syntax. For example, if the formula for the angle of a shape is `"=90[deg,A]"`, the result is displayed as `"90 deg."` if the **DefaultAngleUnits** property is **visDegrees**, and as `"1.5708 rad."` if the **DefaultAngleUnits** property is **visRadians**.

Alternatively, a program can use the following statement to set the cell's result to default angle units.

```vb
vsoCell.Result(visAngleUnits) = 90
```

In this case, the result is 90 degrees if the **DefaultAngleUnits** property is **visDegrees**, and 90 radians if the **DefaultAngleUnits** property is **visRadians**.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]