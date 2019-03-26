---
title: Application.GetOption method (Access)
keywords: vbaac10.chm12503
f1_keywords:
- vbaac10.chm12503
ms.prod: access
api_name:
- Access.Application.GetOption
ms.assetid: 32736ddf-3551-07f5-1559-d0e139c1697d
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.GetOption method (Access)

The **GetOption** method returns the current value of an option in the **Access Options** dialog box, available by choosing the Microsoft Office button, and then choosing **Access Options**. **Variant**.


## Syntax

_expression_.**GetOption** (_OptionName_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _OptionName_|Required|**String**|The name of the option. For a list of option name argument strings, see [Set options from Visual Basic](../access/Concepts/Settings/set-options-from-visual-basic.md).|

## Return value

Variant


## Remarks

The **GetOption** and **[SetOption](access.application.setoption.md)** methods provide a means of changing environment options from Visual Basic code. With these methods, you can set or read any option available in the **Access Options** dialog box.

The available option settings depend on the type of option being set. There are three general types of options:

- Yes/No options that can be set by selecting or clearing a check box.
    
- Options that can be set by entering a string or numeric value.
    
- Predefined options that can be chosen from a list box, combo box, or option group.
    
For options that the user sets by selecting or clearing a check box, the **GetOption** method returns **True** (1) if the option setting is **Yes** (the check box is selected) or **False** (0) if the option setting is **No** (the check box is cleared). To set an option of this kind by using the **SetOption** method, specify **True** or **False** for the setting argument, as in the following example.

```vb
Application.SetOption "Show Status Bar", True
```

<br/>

For options that the user sets by typing a string or numeric value, the **GetOption** method returns the setting as it is displayed in the dialog box. The following example returns a string containing the left margin setting.

```vb
Dim varSetting As Variant 
varSetting = Application.GetOption("Left Margin")
```

<br/>

To set this type of option by using the **SetOption** method, specify the string or numeric value that would be typed in the dialog box. The following example sets the default form template to OrderTemplate.

```vb
Application.SetOption "Form Template", "OrderTemplate"
```

<br/>

For options with settings that are choices in list boxes or combo boxes, the **GetOption** method returns a number corresponding to the position of the setting in the list. Indexing begins with zero, so the **GetOption** method returns zero for the first item, 1 for the second item, and so on. For example, if the **Default Field Type** option on the **Object Designers** tab is set to AutoNumber, the sixth item in the list, the **GetOption** method returns 5.

To set this type of option, specify the option's numeric position within the list as the setting argument for the **SetOption** method. The following example sets the **Default Field Type** option to AutoNumber.

```vb
Application.SetOption "Default Field Type", 5
```

<br/>

Other options are set by choosing an option button in an option group in the **Access Options** dialog box. In Visual Basic, these options are also set by specifying a particular option's position within the option group. The first option in the group is numbered zero, the second, 1, and so on. For example, if the **Selection Behavior** option on the **Object Designers** tab is set to Partially Enclosed, the **GetOption** method returns zero, as in the following example.

```vb
Debug.Print Application.GetOption("Selection Behavior")
```

<br/>

To set an option that's a member of an option group, specify the index number of the option within the group. The following example sets **Selection Behavior** to Fully Enclosed.

```vb
Application.SetOption "Selection Behavior", 1
```

> [!NOTE] 
> When you quit Microsoft Access, you can reset all options to their original settings by using the **SetOption** method on all changed options. You may want to create public variables to store the values of the original settings. You might include code to reset options in the **Close** event procedure for a form, or in a custom exit procedure that the user must run to quit the application.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]