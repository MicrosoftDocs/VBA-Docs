---
title: Add method (Microsoft Forms)
keywords: fm20.chm5224953
f1_keywords:
- fm20.chm5224953
ms.prod: office
ms.assetid: 1030fff9-736c-9ece-5600-2e4f3b4281b8
ms.date: 11/15/2018
localization_priority: Normal
---


# Add method (Microsoft Forms)

Adds or inserts a **Tab** or **Page** in a **[TabStrip](tabstrip-control.md)** or **[MultiPage](multipage-control.md)**, or adds a control by its programmatic identifier (_ProgID_) to a page or form.

## Syntax

For MultiPage, TabStrip:<br/> 
**Set**_Object_ = _object_. **Add(** [ _Name_ [, _Caption_ [, _index_ ]]] **)**

For other controls:<br/> 
**Set**_Control_ = _object_. **Add(**_ProgID_ [, _Name_ [, _Visible_ ]] **)**

<br/>

The **Add** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _Name_|Optional. Specifies the name of the object being added. If a name is not specified, the system generates a default name based on the rules of the application where the form is used.|
| _Caption_|Optional. Specifies the caption to appear on a tab or a control. If a caption is not specified, the system generates a default caption based on the rules of the application where the form is used.|
| _index_|Optional. Identifies the position of a page or tab within a **Pages** or **Tabs** collection. If an index is not specified, the system appends the page or tab to the end of the **Pages** or **Tabs** collection and assigns the appropriate index value.|
| _ProgID_|Required. Programmatic identifier. A text string with no spaces that identifies an object class. The standard syntax for a  _ProgID_ is <Vendor>.<Component>.<Version>. A _ProgID_ is mapped to a class identifier (CLSID).|
| _Visible_|Optional. **True** if the object is visible (default). **False** if the object is hidden.|

## Settings

_ProgID_ values for individual controls are:

|Control|_ProgID_ value|
|:-----|:-----|
|**[CheckBox](checkbox-control.md)**|Forms.CheckBox.1|
|**[ComboBox](combobox-control.md)**|Forms.ComboBox.1|
|**[CommandButton](commandbutton-control.md)**|Forms.CommandButton.1|
|**[Frame](frame-control.md)**|Forms.Frame.1|
|**[Image](image-control.md)**|Forms.Image.1|
|**[Label](label-control.md)**|Forms.Label.1|
|**[ListBox](listbox-control.md)**|Forms.ListBox.1|
|**[MultiPage](multipage-control.md)**|Forms.MultiPage.1|
|**[OptionButton](optionbutton-control.md)**|Forms.OptionButton.1|
|**[ScrollBar](scrollbar-control.md)**|Forms.ScrollBar.1|
|**[SpinButton](spinbutton-control.md)**|Forms.SpinButton.1|
|**[TabStrip](tabstrip-control.md)**|Forms.TabStrip.1|
|**[TextBox](textbox-control.md)**|Forms.TextBox.1|
|**[ToggleButton](togglebutton-control.md)**|Forms.ToggleButton.1|

## Remarks

For a **MultiPage** control, the **Add** method returns a **Page** object. For a **TabStrip**, it returns a **Tab** object. The index value for the first **Page** or **Tab** of a [collection](../../Glossary/vbe-glossary.md#collection) is 0, the value for the second **Page** or **Tab** is 1, and so on.

For the **Controls** collection of an object, the **Add** method returns a control corresponding to the specified _ProgID_. The AddControl event occurs after the control is added.

You can add a control to a user form's **Controls** collection at design time, but you must use the **Designer** property of the Microsoft Visual Basic for Applications Extensibility Library to do so. The **Designer** property returns the **[UserForm](userform-window.md)** object.

The following syntax will return the **Text** property of the specified control:

```vb
userform1.thebox.text

```

If you add a control at run time, you must use the exclamation syntax to reference properties of that control. For example, to return the **Text** property of a control added at run time, use the following syntax:

```vb
userform1!thebox.text

```

> [!NOTE] 
> You can change a control's **Name** property at [run time](../../Glossary/vbe-glossary.md#run-time) only if you added that control at run time with the **Add** method.


## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
