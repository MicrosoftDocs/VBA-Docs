---
title: SpinButton control
keywords: fm20.chm5224986
f1_keywords:
- fm20.chm5224986
ms.prod: office
ms.assetid: 4fca5573-f581-3e1c-55d5-a1e34ec96b04
ms.date: 11/15/2018
localization_priority: Normal
---


# SpinButton control

Increments and decrements numbers.

## Remarks

Clicking a **SpinButton** changes only the value of the **SpinButton**. You can write code that uses the **SpinButton** to update the displayed value of another control. For example, you can use a **SpinButton** to change the month, the day, or the year shown on a date. 

You can also use a **SpinButton** to scroll through a range of values or a list of items, or to change the value displayed in a text box.

To display a value updated by a **SpinButton**, you must assign the value of the **SpinButton** to the displayed portion of a control, such as the **Caption** property of a **[Label](label-control.md)** or the **Text** property of a **[TextBox](textbox-control.md)**. 

To create a horizontal or vertical **SpinButton**, drag the sizing handles of the **SpinButton** horizontally or vertically on the form.

The default property for a **SpinButton** is the **Value** property. The default event for a **SpinButton** is the Change event.

## See also

- [SpinButton object](../../../api/Outlook.spinbutton.object.md)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]