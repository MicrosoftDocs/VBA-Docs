---
title: Creating forms and dialog boxes with right-to-left extensions
keywords: fm20.chm5282667
f1_keywords:
- fm20.chm5282667
ms.prod: office
ms.assetid: 9a36b313-6996-980c-820f-876cb6fdf68d
ms.date: 12/29/2018
localization_priority: Normal
---


# Creating forms and dialog boxes with right-to-left extensions

You can use the Visual Basic Editor and Microsoft Forms version 2.0 in all Microsoft Office 2000 applications to create forms and dialog boxes. Bidirectional extensions to the editor and Microsoft Forms 2.0 are provided in Office 2000 for right-to-left, left-to-right, and mixed-text usage. For a general overview of the standard tools, see the "Forms" topic in Help for the application that you are working in.

Three Microsoft Forms 2.0 properties are generally used to add bidirectional characteristics to forms and dialog boxes. These properties are listed and described in the following table.

|Property|Exposed on|Affects|
|:-----|:-----|:-----|
|[Alignment](../../reference/User-Interface-Help/alignment-property.md)|Controls|Controls|
|[TextAlign](../../reference/User-Interface-Help/textalign-property.md)|Controls|Controls|
|[RightToLeft](../../reference/User-Interface-Help/righttoleft-property-microsoft-forms.md)|Forms|Forms and controls|

<br/>

These properties affect the controls listed in the following table, which are available in the Control Toolbox. You can set these properties in the Properties window in the editor or by using Visual Basic for Applications statements.

|Control|Alignment|TextAlign|RightToLeft|
|:-----|:-----|:-----|:-----|
|CheckBox|**X**|**X**|**X**|
|ComboBox| |**X**|**X**|
|Frame| | |**X**|
|Label| |**X**| |
|ListBox| |**X**|**X**|
|MultiPage| | |**X**|
|OptionButton|**X**|**X**| |
|TabStrip| | |**X**|
|TextBox| |**X**|**X**|
|ToggleButton| |**X**| |


> [!NOTE] 
> Context reading order is generally assigned to text in controls. This means that the reading order of displayed text strings that begin with a non-left-to-right character (for example, text strings in Arabic) will be displayed in right-to-left reading order, and text strings that begin with a left-to-right character will be displayed in left-to-right reading order.


## See also

- [Microsoft Forms reference](../../reference/user-interface-help/reference-microsoft-forms.md)
- [Microsoft Forms conceptual topics](../../reference/user-interface-help/concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]