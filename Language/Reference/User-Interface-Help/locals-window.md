---
title: Locals window
ms.prod: office
ms.assetid: 32e88a9a-853c-e0ec-37ba-364706cf2958
ms.date: 11/21/2018
localization_priority: Normal
---


# Locals window

![Locals window](../../../images/local_ZA01201622.gif)

Automatically displays all of the declared variables in the current procedure and their values.

When the Locals window is visible, it is automatically updated every time there is a change from [run time](../../Glossary/vbe-glossary.md#run-time) to [break mode](../../Glossary/vbe-glossary.md#break-mode) or you navigate in the stack display.

You can:

- Resize the column headers by dragging the border to the right or the left.
    
- Close the window by clicking the **Close** box. If the **Close** box is not visible, double-click the title bar to make the **Close** box visible, and then select it.
    
## Window elements

|Element|Description|
|:------|:----------|
|**Calls Stack** button|Opens the **Call Stack** dialog box, which lists the procedures in the call stack.|
|**Expression**|Lists the name of the variables.<br/><br/>The first variable in the list is a special module variable and can be expanded to display all module level variables in the current module. For a class module, the system variable `<Me>` is defined. For standard modules, the first variable is the `<name of the current module>`. Global variables and variables in other projects are not accessible from the Locals window.<br/><br/>You cannot edit data in this column.|
|**Value**|Lists the value of the variable.<br/><br/>When you click a value in the Value column, the cursor changes to an I-beam. You can edit a value and then press ENTER, the UP ARROW key, the DOWN ARROW key, TAB, SHIFT+TAB, or click on the screen to validate the change. If the value is illegal, the Edit field remains active and the value is highlighted. A message box describing the error also appears. Cancel a change by pressing ESC.<br/><br/>All numeric variables must have a value listed. String variables can have an empty **Value** list.<br/><br/>Variables that contain subvariables can be expanded and collapsed. Collapsed variables do not display a value, but each subvariable does. The expand icon ![Expand icon](../../../images/expand_ZA01201606.gif) and the collapse icon ![Collapse icon](../../../images/collapse_ZA01201589.gif) appear to the left of the variable.|
|**Type**|Lists the variable type. You cannot edit data in this column.|

## See also

- [Window elements](../window-elements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
