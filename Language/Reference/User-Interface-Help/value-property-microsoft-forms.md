---
title: Value property (Microsoft Forms)
keywords: fm20.chm2002180
f1_keywords:
- fm20.chm2002180
ms.prod: office
ms.assetid: bd61f3ae-54b3-6382-6ecf-0c5598279330
ms.date: 11/15/2018
localization_priority: Normal
---


# Value property (Microsoft Forms)

Specifies the state or content of a given control.

## Syntax

_object_.**Value** [= _Variant_ ]

The **Value** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Variant_|Optional. The state or content of the control.|

## Settings

|Control|Description|
|:------|:----------|
|**[CheckBox](checkbox-control.md)**|An integer value indicating whether the item is selected:<br/><br/>Null. Indicates the item is in a null state, neither selected nor [cleared](../../Glossary/glossary-vba.md#clear).<br/><br/>-1 True. Indicates the item is selected.<br/><br/>0 False. Indicates the item is cleared.|
|**[OptionButton](optionbutton-control.md)**|Same as **CheckBox**.|
|**[ToggleButton](togglebutton-control.md)**|Same as **CheckBox**.|
|**[ScrollBar](scrollbar-control.md)**|An integer between the values specified for the **Max** and **Min** properties.|
|**[SpinButton](spinbutton-control.md)**|Same as **ScrollBar**.|
|**[ComboBox](combobox-control.md)**, **[ListBox](listbox-control.md)**|The value in the **BoundColumn** of the currently selected rows.|
|**[CommandButton](commandbutton-control.md)**|Always **False**.|
|**[MultiPage](multipage-control.md)**|An integer indicating the currently active page.<br/><br/>Zero (0) indicates the first page. The maximum value is one less than the number of pages.|
|**[TextBox](textbox-control.md)**|The text in the edit region.|

## Remarks

For a **CommandButton**, setting the **Value** property to **True** in a macro or procedure initiates the button's Click event.

For a **ComboBox**, changing the contents of **Value** does not change the value of **BoundColumn**. To add or delete entries in a **ComboBox**, you can use the **AddItem** or **RemoveItem** method.

**Value** cannot be used with a multi-select list box.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]