---
title: Object Browser (Visual Basic for Applications)
description: Displays the classes, properties, methods, events, and constants available from object libraries and the procedures in your project, along with Help for modules and enumerations.
ms.prod: office
keywords: vbui6.chm181036
f1_keywords:
- vblr6.chm1018983
- vblr6.chm1113295
- vblr6.chm1011390
- vblr6.chm1011391
- vblr6.chm1011392
- vblr6.chm1018936
- vblr6.chm915166
- vblr6.chm1113294
- vblr6.chm1011393
- vblr6.chm1011394
- vblr6.chm1113302
- vblr6.chm1011395
- vblr6.chm1011396
- vblr6.chm1113303
- vblr6.chm1113544
- vblr6.chm1113364
- vblr6.chm1113633
- vblr6.chm1113545
- vblr6.chm1113635
- vblr6.chm1113553
- vblr6.chm1113563
- vblr6.chm1113570
- vblr6.chm1113583
- vblr6.chm1113592
- vblr6.chm1011389
- vblr6.chm1113608
- vblr6.chm1113636
- vblr6.chm1113634
- vblr6.chm1113637
- vbui6.chm181036
ms.assetid: bee7d672-7cb5-2cd7-86a2-00110d40c6bc
ms.date: 12/17/2018
localization_priority: Normal
---


# Object Browser

![Object browser](../../../images/objbrows_ZA01201634.gif)

A dialog box in which you can examine the contents of an object library to get information about the objects provided.

Displays the [classes](../../Glossary/vbe-glossary.md#class), properties, methods, events, and constants available from [object libraries](../../Glossary/vbe-glossary.md#object-library) and the [procedures](../../Glossary/vbe-glossary.md#procedure) in your project. You can use it to find and use objects you create, as well as objects from other applications.

You can get Help for the **Object Browser** by searching for Object Browser in Help.


## Window elements 

|Element|Icon|Description|
|:------|:---|:----------|
|**Project/Library** box| |Displays the currently referenced libraries for the active project. You can add libraries in the **References** dialog box. `<All Libraries>` allows all of the libraries to be displayed at one time.|
|**Search Text** box | |Contains the string that you want to use in your search. You can type or choose the string you want. The **Search Text** box contains the last four search strings that you enter until you close the project. You can use the standard Visual Basic wildcards when typing a string.<br/><br/>If you want to search for a whole word, you can use the **Find Whole Word Only** command on the shortcut menu.|
|**Go Back** button| ![Go Back button](../../../images/goback_ZA01201613.gif) |Allows you to go back to the previous selection in the **Classes** and **Members Of** lists. Each time you click it, you move back one selection until all of your choices are exhausted.|
|**Go Forward** button| ![Go Forward button](../../../images/forward_ZA01201610.gif)| Allows you to repeat your original selections in the **Classes** and **Members Of** lists each time you click it, until you exhaust the list of selections.|
|**Copy to Clipboard** button| ![Copy to Clipboard button](../../../images/but_copy_ZA01201582.gif)| Copies the current selection in the **Members Of** list or the **Details** pane text to the Clipboard. You can then paste the selection into your code.|
|**View Definition** button|![View Definition button](../../../images/viewdef_ZA01201805.gif)|  Moves the cursor to the place in the Code window where the selection in the **Members Of** list or the **Classes** list is defined.|
|**Help** button| ![Help button](../../../images/but_help_ZA01201583.gif) |Displays the online Help topic for the item selected in the **Classes** and **Members Of** lists. You can also use F1.|
|**Search** button| ![Search button](../../../images/search_ZA01201651.gif) |Initiates a search of the libraries for the class or property, method, event, or constant that matches the string you typed in the **Search Text** box, and opens the **Search Results** pane with the appropriate list of information.|
|**Show/Hide Search Results** button| ![Show/Hide Search Results button](../../../images/showsear_ZA01201652.gif) |Opens or hides the **Search Results** pane. The **Search Results** pane changes to show the search results from the project or library chosen in the **Project/Library** list. Search results are listed alphabetically from A to Z.|
|**Search Results** list| |Displays the library, class, and member that corresponds to the items that contain your search string. The **Search Results** pane changes when you change the selection in the **Project/Library** box.|
|**Classes** list| |Displays all of the available classes in the library or project selected in the **Project/Library** box. If there is code written for a class, that class appears in bold. The list always begins with `<globals>`, a list of globally accessible members.<br/><br/>If you select a class and do not specify a member, you will get the default member if one is available. The default member is identified by an asterisk (*) or by the default [icon](icons-used-in-the-object-browser-and-code-windows.md) specific to the member.|
|**Members Of** list| |Displays the elements of the class selected in the **Classes** pane by group and then alphabetically within each group. Methods, properties, events, or constants that have code written for them appear bold. You can change the order of this list with the **Group Members** command on the **Object Browser** shortcut menu.|
|**Details** pane| |Shows the definition of the member. The **Details** pane contains a jump to the class or library to which the element belongs. Some members have jumps to their parent class.<br/><br/>For example, if the text in the **Details** pane states that Command1 is declared as a command button type, clicking the command button takes you to the **Command Button** class.<br/><br/>You can copy or drag text from the **Details** pane to the Code window.|
|**Split** bar| |Splits the panes so that you can adjust their size. There are splits between the:<br/><br/>- **Classes** box and the **Members Of** box.<br/>- **Search Results** list and the **Classes** and **Members Of** boxes.<br/>- **Classes** and **Members Of** boxes and the **Details** pane.|

## Modules

|Module|Description|To get Help on a particular property, procedure, or constant|
|:-----|:----------|:-----------------------------------------------------------|
|**Collection**| Contains procedures used to perform operations on the **Collection** object. These constants can be used anywhere in your code.|1. Select the procedure from the **Members of 'Collection'** list.<br/>2. Choose the **Help** ![Help button](../../../images/but_help_ZA01201583.gif) button. |
|**ColorConstants**| Contains predefined color constants. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'ColorConstants'** list.<br/>2. Choose the **Help** button. |
|**Constants**| Contains miscellaneous constants. These constants can be used anywhere in your code.| 1. Select the constant from the **Members of 'Constants'** list.<br/>2. Choose the **Help** button. |
|**Conversion**| Contains the procedures used to perform various conversion operations. These constants can be used anywhere in your code.|1. Select the procedure from the **Members of 'Conversion'** list.<br/>2. Choose the **Help** button.<br/><br/>**NOTE**: When you use **Variant** variables, explicit data-type conversions are unnecessary. |
|**DateTime**| Contains the procedures and properties used in date and time operations. These constants can be used anywhere in your code.|1. Select the procedure from the **Members of 'DateTime'** list.<br/>2. Choose the **Help** button. |
|**ErrObject**| Contains properties and procedures used to identify and handle run-time errors by using the **Err** object. These constants can be used anywhere in your code.|1. Select the property or procedure from the **Members of 'ErrObject'** list.<br/>2. Choose the **Help** button. |
|**FileSystem**| Contains the procedures used to perform file, directory or folder, and system operations. These constants can be used anywhere in your code.|1. Select the procedure from the **Members of 'FileSystem'** list.<br/>2. Choose the **Help** button. |
|**Financial**| Contains procedures used to perform financial operations. These constants can be used anywhere in your code.|1. Select the procedure from the **Members of 'Financial'** list.<br/>2. Choose the **Help** button. |
|**FormShowConstants**| Contains predefined **[UserForm](userform-window.md)** constants. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'FormShowConstants'** list.<br/>2. Choose the **Help** button. |
|**Global**| Contains procedures and properties used to perform operations on the **[UserForm](userform-object.md)** object. These constants can be used anywhere in your code.|1. Select the procedure from the **Members Of 'Global'** list.<br/>2. Choose the **Help** button. |
|**Information**| Contains the procedures used to return, test for, or verify information. These constants can be used anywhere in your code.|1. Select the procedure from the **Members of 'Information'** list.<br/>2. Choose the **Help** button. |
|**Interaction**| Contains procedures used to interact with objects, applications, and systems. These constants can be used anywhere in your code.|1. Select the procedure from the **Members of 'Interaction'** list.<br/>2. Choose the **Help** button. |
|**KeyCodeConstants** | Contains predefined keycode constants that can be used anywhere in your code.|1. Select the constant from the **Members of 'KeyCodeConstants'** list.<br/>2. Choose the **Help** button. |
|**Math**| Contains procedures used to perform mathematical operations. These constants can be used anywhere in your code.|1. Select the procedure from the **Members of 'Math'** list.<br/>2. Choose the **Help** button. |
|**String**| Contains procedures used to perform string operations. These constants can be used anywhere in your code.|1. Select the procedure from the **Members of 'Strings'** list.<br/>2. Choose the **Help** button. |
|**SystemColorConstants**|Contains constants that identify various parts of the graphical user interface. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'SystemColorConstants'** list.<br/>2. Choose the **Help** button. |


## Enumerations

|Enumeration|Description|To get Help on a particular constant|
|:----------|:----------|:-----------------------------------|
|**VbAppWinStyle**| Contains constants used by the **Shell** function to control the style of an application window. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbAppWinStyle'** list.<br/>2. Choose the **Help** ![Help button](../../../images/but_help_ZA01201583.gif) button. |
|**VbCalendar**| Contains constants used to determine the type of calendar used by Visual Basic. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbCalendar'** list.<br/>2. Choose the **Help** button. |
|**VbCallType**| Defines constants used to identify the call type used by the **CallByName** function.|1. Select the constant from the **Members of 'VbCallType'** list.<br/>2. Choose the **Help** button. |
|**VbCompareMethod**| Contains constants used to determine the way strings are compared when using the **Instr** and **StrComp** functions. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbCompareMethod'** list.<br/>2. Choose the **Help** button. |
|**VbDateTimeFormat**| Defines constants used to identify how the date and time are formatted.|1. Select the constant from the **Members of 'VbDateTimeFormat'** list.<br/>2. Choose the **Help** button. |
|**VbDayOfWeek**| Contains constants used to identify specific days of the week when using the **DateDiff**, **DatePart**, and **Weekday** functions. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbDayOfWeek'** list.<br/>2. Choose the **Help** button. |
|**VbFileAttribute**| Contains constants used to identify file attributes used in the **Dir**, **GetAttr**, and **SetAttr** functions. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbFileAttribute'** list.<br/>2. Choose the **Help** button. |
|**VbFirstWeekOfYear**| Contains constants used to identify how the first week of a year is determined when using the **DateDiff** and **DatePart** functions. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbFirstWeekOfYear'** list.<br/>2. Choose the **Help** button. |
|**VbIMEStatus**| Available only in East Asia versions, contains constants used to identify the Input Method Editor (IME) when using the **IMEStatus** function. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbIMEStatus'** list.<br/>2. Choose the **Help** button. |
|**VbMsgBoxResult**| Contains constants used to identify which button was pressed on a message box displayed by using the **MsgBox** function. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbMsgBoxResult'** list.<br/>2. Choose the **Help** button. |
|**VbMsgBoxStyle**| Contains constants used to specify the behavior of a message box, along with symbols and buttons that appear on it, when displayed by using the **MsgBox** function. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbMsgBoxStyle'** list.<br/>2. Choose the **Help** button. |
|**VbQueryClose**| Contains constants used to identify what caused the **QueryClose** event to occur. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbQueryClose'** list.<br/>2. Choose the **Help** button. |
|**VbStrConv**| Contains constants used to identify the type of string conversion to be performed by the **StrConv** function. These constants can be used anywhere in your code.|1. Select the constant from the **Members of 'VbStrConv'** list.<br/>2. Choose the **Help** button. |
|**VbTriState**| Defines constants used to identify one of three possible states.|1. Select the constant from the **Members of 'VbTriState'** list.<br/>2. Choose the **Help** button. |
|**VbVarType**| Contains constants used to identify the various types of data that can be contained in a **Variant**. These constants match the return values of the **VarType** function and can be used anywhere in your code.|1. Select the constant from the **Members of 'VbVarType'** list.<br/>2. Choose the **Help** button. |

    
## See also

- [Object Browser shortcut menu](shortcut-menu.md#object-browser)
- [Use the Object Browser](use-the-object-browser.md)
- [Code window and Object Browser icons](icons-used-in-the-object-browser-and-code-windows.md)
- [Object Browser on the View menu](view-menu.md)
- [Window elements](../window-elements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]