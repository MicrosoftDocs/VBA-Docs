---
title: ListBox.LabelY property (Access)
keywords: vbaac10.chm11269
f1_keywords:
- vbaac10.chm11269
ms.prod: access
api_name:
- Access.ListBox.LabelY
ms.assetid: 79a1486b-4f51-fabd-e56e-51cb2868c0c2
ms.date: 02/22/2019
localization_priority: Normal
---


# ListBox.LabelY property (Access)

The **LabelY** property (along with the **LabelX** property) specifies the placement of the label for a new control. Read/write **Integer**.



## Syntax

_expression_.**LabelY**

_expression_ A variable that represents a **[ListBox](Access.ListBox.md)** object.


## Remarks

If the orientation is left to right for a form or report, **LabelX** and **LabelY** behavior matches standard Microsoft Access left-to-right orientation. For more information about orientation, see the **[Orientation](access.form.orientation.md)** property.

If the orientation is right to left, the origin of the coordinate system for **LabelX** and **LabelY** is the upper-right corner of the attached control. A negative number for **LabelX** places the label to the right of the control. A negative number for **LabelY** places the label above the control.

For General and Right alignment when the orientation is right to left, **LabelX** and **LabelY** specify the location of the upper-right corner of the label relative to the upper-right corner of the label's attached control. 

For Left and Center alignment, **LabelX** and **LabelY** specify the location of the upper-left corner and top center, respectively, of the label relative to the upper-right corner of the label's attached control.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

