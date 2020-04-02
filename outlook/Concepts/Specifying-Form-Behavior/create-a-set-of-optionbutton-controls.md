---
title: Create a Set of OptionButton Controls
ms.prod: outlook
ms.assetid: 6aee3c64-df73-df1a-0db8-2740f8dec0d9
ms.date: 06/08/2019
localization_priority: Normal
---


# Create a Set of OptionButton Controls

By default, all [OptionButton](../../../api/Outlook.optionbutton.md)controls in a container are part of a single option group. This means that selecting one of the buttons automatically sets all other option buttons on the form to **False**.

If you want more than one option group on the form, there are two ways to create additional groups:

- Use the [GroupName](../../../api/Outlook.optionbutton.groupname.md)property to identify related buttons. This method reduces the number of controls required on the form, which can reduce the hard disk space required and improve the performance of the form. If you want to create an option group in a [TabStrip](../../../api/Outlook.tabstrip.md)(which is not a container), you must use the **GroupName** property. For more information, see [How to: Create a Set of OptionButtons by Using the GroupName Property](create-a-set-of-optionbuttons-by-using-the-groupname-property.md).
    
- Put related buttons in a **[Page](../../../api/Outlook.page.md)**, **[MultiPage](../../../api/Outlook.multipage.md)**, or **[Frame](../../../api/Outlook.frame.md)** on the form. For more information, see [How to: Add a Control to a Form](add-a-control-to-a-form.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]