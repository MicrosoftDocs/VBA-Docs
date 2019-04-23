---
title: NavigationButton.LabelY property (Access)
keywords: vbaac10.chm10485
f1_keywords:
- vbaac10.chm10485
ms.prod: access
api_name:
- Access.NavigationButton.LabelY
ms.assetid: 09adfd0d-c877-564a-0ddb-a2b3caf5b085
ms.date: 02/22/2019
localization_priority: Normal
---


# NavigationButton.LabelY property (Access)

The **LabelY** property (along with the **LabelX** property) specifies the placement of the label for a new control. Read/write **Integer**.



## Syntax

_expression_.**LabelY**

_expression_ A variable that represents a **[NavigationButton](Access.NavigationButton.md)** object.


## Remarks

If the orientation is left to right for a form or report, **LabelX** and **LabelY** behavior matches standard Microsoft Access left-to-right orientation. For more information about orientation, see the **[Orientation](access.form.orientation.md)** property.

If the orientation is right to left, the origin of the coordinate system for **LabelX** and **LabelY** is the upper-right corner of the attached control. A negative number for **LabelX** places the label to the right of the control. A negative number for **LabelY** places the label above the control.

For General and Right alignment when the orientation is right to left, **LabelX** and **LabelY** specify the location of the upper-right corner of the label relative to the upper-right corner of the label's attached control. 

For Left and Center alignment, **LabelX** and **LabelY** specify the location of the upper-left corner and top center, respectively, of the label relative to the upper-right corner of the label's attached control.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

