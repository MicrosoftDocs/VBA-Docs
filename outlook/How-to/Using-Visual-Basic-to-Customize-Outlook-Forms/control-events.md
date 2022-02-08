---
title: Control Events
keywords: olfm10.chm3077124
f1_keywords:
- olfm10.chm3077124
ms.prod: outlook
ms.assetid: 6305af2d-d26c-024f-945a-8eaa773bab85
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Control Events

Most Microsoft Forms 2.0 controls in an Microsoft Outlook custom form support only one event, the **Click** event.
A control bound to a field does not fire the **Click** event. You must handle the appropriate [field event](field-events.md) to detect a user's interaction with a control bound to a field.
The following Forms 2.0 controls and Outlook controls fire the **Click** event whenever a user clicks anywhere in the control.
 **[CheckBox](../../../api/Outlook.checkbox.md)**
 **[CommandButton](../../../api/Outlook.commandbutton.md)**
 **[Frame](../../../api/Outlook.frame.md)**
 **[Image](../../../api/Outlook.image.md)**
 **[Label](../../../api/Outlook.label.md)**
 **[OptionButton](../../../api/Outlook.optionbutton.md)**
 **[ToggleButton](../../../api/Outlook.togglebutton.md)**
 **[OlkBusinessCardControl](../../../api/Outlook.OlkBusinessCardControl.md)**
 **[OlkCategory](../../../api/Outlook.OlkCategory.md)**
 **[OlkCheckBox](../../../api/Outlook.OlkCheckBox.md)**
 **[OlkCommandButton](../../../api/Outlook.OlkCommandButton.md)**
 **[OlkContactPhoto](../../../api/Outlook.OlkContactPhoto.md)**
 **[OlkDateControl](../../../api/Outlook.OlkDateControl.md)**
 **[OlkFrameHeader](../../../api/Outlook.OlkFrameHeader.md)**
 **[OlkInfoBar](../../../api/Outlook.OlkInfoBar.md)**
 **[OlkLabel](../../../api/Outlook.OlkLabel.md)**
 **[OlkOptionButton](../../../api/Outlook.OlkOptionButton.md)**
 **[OlkSenderPhoto](../../../api/Outlook.OlkSenderPhoto.md)**
 **[OlkTextBox](../../../api/Outlook.OlkTextBox.md)**
 **[OlkTimeControl](../../../api/Outlook.OlkTimeControl.md)**
 **[OlkTimeZoneControl](../../../api/Outlook.OlkTimeZoneControl.md)**

The following controls fire the **Click** event when the user selects an item in the list.
 **[ComboBox](../../../api/Outlook.combobox.md)**
 **[ListBox](../../../api/Outlook.listbox.md)**
 **[OlkComboBox](../../../api/Outlook.OlkComboBox.md)**
 **[OlkListBox](../../../api/Outlook.OlkListBox.md)**

The following controls don't support the **Click** event.
 **[ScrollBar](../../../api/Outlook.scrollbar.md)**
 **[SpinButton](../../../api/Outlook.spinbutton.md)**
 **[TabStrip](../../../api/Outlook.tabstrip.md)**
 **[TextBox](../../../api/Outlook.textbox.md)**

While the **MultiPage** control itself does not support the **Click** event, an individual **[Page](../../../api/Outlook.page.md)** on a **MultiPage** control will fire the **Click** event if the user clicks inside the client area of the page, but not if the user clicks the tab associated with the page.

To detect a change in a **TextBox** control, bind the control to a field and then handle the appropriate field event.
If you have to further extend controls in a custom form, customize a form with Outlook controls in a form region instead of Forms 2.0 controls in a form page. For more information, see [Controls in a Custom Form](../../Concepts/Forms/controls-in-a-custom-form.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]