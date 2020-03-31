---
title: Create a shortcut menu for a form, form control, or report
ms.prod: access
ms.assetid: 56fe8923-053f-e04d-78d6-c4dd814b6499
ROBOTS: INDEX
ms.date: 06/08/2019
localization_priority: Normal
---


# Create a shortcut menu for a form, form control, or report

**Applies to**: Access 2013

When you are designing a form or report, you may want to provide a method for a user to easily use a command that applies only to the current context. One way to do this is to create a custom shortcut menu and apply it to a form report, or control. The shortcut menu appears when the user right-clicks the object to which the shortcut menu is applied.

In earlier versions of **Access**, you could use the **Customize** dialog box to create custom shortcut menus. In Access 2013, you must use Visual Basic for Applications (VBA) code to create a shortcut menu. This article describes you how to create a shortcut menu using VBA.

To create a shortcut menu, you first have to create a **[CommandBar](../../../api/Office.CommandBar.md)** object. The **CommandBar** object represents the shortcut menu. Then, you use the **[Add](../../../api/Office.CommandBarControls.Add.md)** method to create **[CommandBarControl](../../../api/Office.CommandBarControl.md)** objects. Each time that you create a CommandBarControl object, a command is added to the shortcut menu.

The following example creates a shortcut menu named **SimpleShortcutMenu** that contains two commands, **Remove Filter/Sort** and **Filter by Selection**.

> [!NOTE] 
> To use the following examples, you must set a reference to the **Microsoft Office 15.0 Object Library**. See [Set References to Type Libraries](../Settings/set-references-to-type-libraries.md) for more information about how to set references.

```vb
Sub CreateSimpleShortcutMenu() 
    Dim cmbShortcutMenu As Office.CommandBar 
     
    ' Create a shortcut menu named "SimpleShortcutMenu. 
    Set cmbShortcutMenu = CommandBars.Add("SimpleShortcutMenu", msoBarPopup, False, True) 
     
    ' Add the Remove Filter/Sort command. 
    cmbShortcutMenu.Controls.Add Type:=msoControlButton, Id:=605 
 
    ' Add the Filter By Selection command. 
    cmbShortcutMenu.Controls.Add Type:=msoControlButton, Id:=640 
     
    Set cmbShortcutMenu = Nothing 
     
End Sub
```

After you've run the code, the shortcut menu is saved as part of the database. You don't have to run the same code to re-create the shortcut menu every time that you open the database.

To assign the shortcut menu to a form, form control, or report, set the **Shortcut Menu** property of the object to **Yes** and set the **Shortcut Menu Bar** property of the object to the name of the shortcut menu. For this example, set the **Shortcut Menu Bar** property to **SimpleShortcutMenu**.

The following example creates a shortcut menu named **cmdFormFiltering** that contains commands that are useful to use with Continuous forms. In this example, the **BeginGroup** property is used on several controls to group controls visually.

```vb
Sub CreateShortcutMenuWithGroups() 
    Dim cmbRightClick As Office.CommandBar 
 
 ' Create the shortcut menu. 
    Set cmbRightClick = CommandBars.Add("cmdFormFiltering", msoBarPopup, False, True) 
     
    With cmbRightClick 
        ' Add the Find command. 
        .Controls.Add msoControlButton, 141, , , True 
         
        ' Start a new grouping and add the Sort Ascending command. 
        .Controls.Add(msoControlButton, 210, , , True).BeginGroup = True 
         
        ' Add the Sort Descending command. 
        .Controls.Add msoControlButton, 211, , , True 
         
        ' Start a new grouping and add the Remove Filer/Sort command. 
        .Controls.Add(msoControlButton, 605, , , True).BeginGroup = True 
         
        ' Add the Filter by Selection command. 
        .Controls.Add msoControlButton, 640, , , True 
         
        ' Add the Filter Excluding Selection command. 
        .Controls.Add msoControlButton, 3017, , , True 
         
        ' Add the Between... command. 
        .Controls.Add msoControlButton, 10062, , , True 
    End With 
 
Set cmbRightClick = Nothing 
End Sub
```

<br/>

The following example creates a shortcut menu named **cmdReportRightClick** that contains commands that are useful to use with a report. This example shows how to change the **Caption** property of each control as they're added to the shortcut menu.

```vb
Sub CreateReportShortcutMenu() 
    Dim cmbRightClick As Office.CommandBar 
    Dim cmbControl As Office.CommandBarControl 
 
   ' Create the shortcut menu. 
    Set cmbRightClick = CommandBars.Add("cmdReportRightClick", msoBarPopup, False, True) 
 
    With cmbRightClick 
         
        ' Add the Print command. 
        Set cmbControl = .Controls.Add(msoControlButton, 2521, , , True) 
        ' Change the caption displayed for the control. 
        cmbControl.Caption = "Quick Print" 
         
        ' Add the Print command. 
        Set cmbControl = .Controls.Add(msoControlButton, 15948, , , True) 
        ' Change the caption displayed for the control. 
        cmbControl.Caption = "Select Pages" 
         
        ' Add the Page Setup... command. 
        Set cmbControl = .Controls.Add(msoControlButton, 247, , , True) 
        ' Change the caption displayed for the control. 
        cmbControl.Caption = "Page Setup" 
         
        ' Add the Mail Recipient (as Attachment)... command. 
        Set cmbControl = .Controls.Add(msoControlButton, 2188, , , True) 
        ' Start a new group. 
        cmbControl.BeginGroup = True 
        ' Change the caption displayed for the control. 
        cmbControl.Caption = "Email Report as an Attachment" 
         
        ' Add the PDF or XPS command. 
        Set cmbControl = .Controls.Add(msoControlButton, 12499, , , True) 
        ' Change the caption displayed for the control. 
        cmbControl.Caption = "Save as PDF/XPS" 
         
        ' Add the Close command. 
        Set cmbControl = .Controls.Add(msoControlButton, 923, , , True) 
        ' Start a new group. 
        cmbControl.BeginGroup = True 
        ' Change the caption displayed for the control. 
        cmbControl.Caption = "Close Report" 
    End With 
     
    Set cmbControl = Nothing 
    Set cmbRightClick = Nothing 
End Sub
```

<a name="AboutContributors"> </a>

## About the contributors

Sample code provided by Edwin Blancovitch, president of [Advanced Developers.net](https://advdev.net/), creators of [Easy Payroll](https://www.easypayroll.net/), a software package to manage your human resources, payroll, scheduling, time, and attendance needs.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
