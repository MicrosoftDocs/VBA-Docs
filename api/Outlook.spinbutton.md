---
title: SpinButton Object (Outlook Forms Script)
keywords: olfm10.chm2000630
f1_keywords:
- olfm10.chm2000630
ms.prod: outlook
ms.assetid: 3221b356-1e68-9e14-48ab-4a30c38aa685
ms.date: 06/08/2017
localization_priority: Normal
---


# SpinButton Object (Outlook Forms Script)

Increments and decrements a value.


## Remarks

Clicking a  **SpinButton** changes only the value of the **SpinButton**. You can write code that uses the  **SpinButton** to update the displayed value of another control. For example, you can use a **SpinButton** to change the month, the day, or the year shown on a date. You can also use a **SpinButton** to scroll through a range of values or a list of items, or to change the value displayed in a text box.

To display a value updated by a  **SpinButton**, you must assign the value of the  **SpinButton** to the displayed portion of a control, such as the **[Caption](Outlook.label.caption.md)** property of a **[Label](Outlook.label.md)** or the **[Text](Outlook.textbox.text.md)** property of a **[TextBox](Outlook.textbox.md)**. To create a horizontal or vertical  **SpinButton**, drag the sizing handles of the  **SpinButton** horizontally or vertically on the form.

The default property for a  **SpinButton** is the **[Value](Outlook.textbox.value.md)** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]