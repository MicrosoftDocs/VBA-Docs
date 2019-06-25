---
title: Application.DefaultDurationUnits property (Visio)
keywords: vis_sdr.chm10051045
f1_keywords:
- vis_sdr.chm10051045
ms.prod: visio
api_name:
- Visio.Application.DefaultDurationUnits
ms.assetid: 11810de2-0c2f-a498-6b7a-090d5397066b
ms.date: 06/25/2019
localization_priority: Normal
---


# Application.DefaultDurationUnits property (Visio)

Determines the default unit of measure for quantities that represent durations. Read/write.


## Syntax

_expression_.**DefaultDurationUnits**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Return value

Variant


## Remarks

The **DefaultDurationUnits** property corresponds to the value shown in the **Duration** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box (**File** tab > **Options**).

The return value contains one of the values of **[VisUnitCodes](Visio.visunitcodes.md)**, which are declared in the Microsoft Visio type library.

You can specify **DefaultDurationUnits** as an integer (a member of **VisUnitCodes**) or a string value such as "minutes". If the string is invalid or the unit code is inappropriate (non-duration), an error is generated.

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About units of measure](../visio/Concepts/about-units-of-measure-visio.md).

Cell formulas that contain a specific unit of measure are displayed in those units regardless of the default duration units setting. Many cell formulas, however, use implicit unit syntax and are displayed in default units.

A program can create a cell whose formula displays in default units by setting the cell's **Formula** property to a string in implicit unit syntax. For example, if a formula specifying duration is `"=10[em,E]"`, the result displays as `"0.0069 ed"` if the **DefaultDurationUnits** property is **visElapsedDay**, and as `"600.0000 es"` if the **DefaultDurationUnits** property is **visElapsedSec**.

Alternatively, a program can use the following statement to set the cell's result to default duration units. 

```vb
vsoCell.Result(visDurationUnits) = 60
```

In this case, the result is 60 minutes if the **DefaultDurationUnits** property is **visElapsedMin** and 60 seconds if the **DefaultDurationUnits** property is **visElapsedSec**.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]