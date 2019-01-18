---
title: Control Events
keywords: olfm10.chm3077124
f1_keywords:
- olfm10.chm3077124
ms.prod: outlook
ms.assetid: 6305af2d-d26c-024f-945a-8eaa773bab85
ms.date: 06/08/2017
localization_priority: Normal
---


# Control Events



Most Microsoft Forms 2.0 controls in an Microsoft Outlook custom form support only one event, the  **Click** event.
A control bound to a field does not fire the  **Click** event. You must handle the appropriate [field event](field-events.md) to detect a user's interaction with a control bound to a field.
The following Forms 2.0 controls and Outlook controls fire the  **Click** event whenever a user clicks anywhere in the control.<br>
 **[CheckBox](../../../api/Outlook.checkbox.md)**<br>
 **[CommandButton](../../../api/Outlook.commandbutton.md)**<br>
 **[Frame](../../../api/Outlook.frame.md)**<br>
 **[Image](../../../api/Outlook.image.md)**<br>
 **[Label](../../../api/Outlook.label.md)**<br>
 **[OptionButton](../../../api/Outlook.optionbutton.md)**<br>
 **[ToggleButton](../../../api/Outlook.togglebutton.md)**<br>
 **[OlkBusinessCardControl](../../../api/Outlook.OlkBusinessCardControl.md)**<br>
 **[OlkCategory](../../../api/Outlook.OlkCategory.md)**<br>
 **[OlkCheckBox](../../../api/Outlook.OlkCheckBox.md)**<br>
 **[OlkCommandButton](../../../api/Outlook.OlkCommandButton.md)**<br>
 **[OlkContactPhoto](../../../api/Outlook.OlkContactPhoto.md)**<br>
 **[OlkDateControl](../../../api/Outlook.OlkDateControl.md)**<br>
 **[OlkFrameHeader](../../../api/Outlook.OlkFrameHeader.md)**<br>
 **[OlkInfoBar](../../../api/Outlook.OlkInfoBar.md)**<br>
 **[OlkLabel](../../../api/Outlook.OlkLabel.md)**<br>
 **[OlkOptionButton](../../../api/Outlook.OlkOptionButton.md)**<br>
 **[OlkSenderPhoto](../../../api/Outlook.OlkSenderPhoto.md)**<br>
 **[OlkTextBox](../../../api/Outlook.OlkTextBox.md)**<br>
 **[OlkTimeControl](../../../api/Outlook.OlkTimeControl.md)**<br>
 **[OlkTimeZoneControl](../../../api/Outlook.OlkTimeZoneControl.md)**<br>
 
The following controls fire the  **Click** event when the user selects an item in the list.<br>
 **[ComboBox](../../../api/Outlook.combobox.md)**<br>
 **[ListBox](../../../api/Outlook.listbox.md)**<br>
 **[OlkComboBox](../../../api/Outlook.OlkComboBox.md)**<br>
 **[OlkListBox](../../../api/Outlook.OlkListBox.md)**<br>

The following controls do not support the  **Click** event.<br>
 **[ScrollBar](../../../api/Outlook.scrollbar.md)**<br>
 **[SpinButton](../../../api/Outlook.spinbutton.md)**<br>
 **[TabStrip](../../../api/Outlook.tabstrip.md)**<br>
 **[TextBox](../../../api/Outlook.textbox.md)**<br>

While the  **MultiPage** control itself does not support the **Click** event, an individual **[Page](../../../api/Outlook.page.md)** on a **MultiPage** control will fire the **Click** event if the user clicks inside the client area of the page, but not if the user clicks the tab associated with the page.<br>

To detect a change in a  **TextBox** control, bind the control to a field and then handle the appropriate field event.
If you have to further extend controls in a custom form, customize a form with Outlook controls in a form region instead of Forms 2.0 controls in a form page. For more information, see  [Controls in a Custom Form](../../Concepts/Forms/controls-in-a-custom-form.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]