---
title: Layout event
keywords: fm20.chm5224946
f1_keywords:
- fm20.chm5224946
ms.prod: office
api_name:
- Office.Layout
ms.assetid: ae4e356a-3283-e6a0-ac29-25327ff7c3df
ms.date: 11/15/2018
localization_priority: Normal
---


# Layout event

Occurs when a form, **[Frame](frame-control.md)**, or **[MultiPage](multipage-control.md)** changes size.

## Syntax

For MultiPage <br/>
**Private Sub**_object_ _**Layout(**_index_**As Long)**

For all other controls <br/>
**Private Sub**_object_ _**Layout( )**

The **Layout** event syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _index_|Required. The index of the page in a **MultiPage** that changed size.|

## Remarks

The default action of the layout event is to calculate new positions of controls and to repaint the screen. A user can initiate the Layout event by changing the size of a control.

For controls that support the **AutoSize** property, the Layout event is initiated when **AutoSize** changes the size of the control. This occurs when the user changes the value of a property that affects the size of a control. For example, increasing the **Font** size of a **[TextBox](textbox-control.md)** or **[Label](label-control.md)** can significantly change the dimensions of the control and initiate a Layout event.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]