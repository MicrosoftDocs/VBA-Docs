---
title: Style property
ms.prod: office
api_name:
- Office.Style
ms.assetid: b951714c-360e-47c7-85a4-c3260d898b1d
ms.date: 11/16/2018
localization_priority: Normal
---


# Style property

For **[ComboBox](combobox-control.md)**, specifies how the user can choose or set the control's value. 

For **[MultiPage](multipage-control.md)** and **[TabStrip](tabstrip-control.md)**, identifies the style of the tabs on the control.

## Syntax

For ComboBox: _object_.**Style** [= _fmStyle_ ]

For MultiPage and TabStrip: _object_.**Style** [= _fmTabStyle_ ]

The **Style** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmStyle_|Optional. Specifies how a user sets the value of a **ComboBox**.|
| _fmTabStyle_|Optional. Specifies the tab style in a **MultiPage** or **TabStrip**.|

## Settings

The settings for _fmStyle_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmStyleDropDownCombo_|0|The **ComboBox** behaves as a drop-down combo box. The user can type a value in the edit region or select a value from the drop-down list (default).|
| _fmStyleDropDownList_|2|The **ComboBox** behaves as a list box. The user must choose a value from the list.|

<br/>

The settings for  _fmTabStyle_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmTabStyleTabs_|0|Displays tabs on the tab bar (default).|
| _fmTabStyleButtons_|1|Displays buttons on the tab bar.|
| _fmTabStyleNone_|2|Does not display the tab bar.|

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]