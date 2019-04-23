---
title: DoCmd object (Access)
keywords: vbaac10.chm4241
f1_keywords:
- vbaac10.chm4241
ms.prod: access
api_name:
- Access.DoCmd
ms.assetid: 3ce44cca-9979-0a1e-9787-079a52ce528f
ms.date: 03/06/2019
localization_priority: Priority
---


# DoCmd object (Access)

You can use the methods of the **DoCmd** object to run Microsoft Office Access actions from Visual Basic. An action performs tasks such as closing windows, opening forms, and setting the value of controls.


## Remarks

For example, you can use the **OpenForm** method of the **DoCmd** object to open a form, or use the **Hourglass** method to change the mouse pointer to an hourglass icon.

Most of the methods of the **DoCmd** object have arguments; some are required, while others are optional. If you omit optional arguments, the arguments assume the default values for the particular method. For example, the **OpenForm** method uses seven arguments, but only the first argument, _FormName_, is required. 

The following example shows how you can open the **Employees** form in the current database. Only employees with the title Sales Representative are included.

```vb
DoCmd.OpenForm "Employees", , ,"[Title] = 'Sales Representative'"
```

The **DoCmd** object doesn't support methods corresponding to the following actions:
    
- MsgBox. Use the **MsgBox** function.    
- RunApp. Use the **Shell** function to run another application.    
- RunCode. Run the function directly in Visual Basic.   
- SendKeys. Use the **SendKeys** statement.   
- SetValue. Set the value directly in Visual Basic.   
- StopAllMacros.   
- StopMacro.
    

## Example

The following example opens a form in Form view and moves to a new record.

```vb
Sub ShowNewRecord() 
 DoCmd.OpenForm "Employees", acNormal 
 DoCmd.GoToRecord , , acNewRec 
End Sub
```


## Methods

- [AddMenu](Access.DoCmd.AddMenu.md)
- [ApplyFilter](Access.DoCmd.ApplyFilter.md)
- [Beep](Access.DoCmd.Beep.md)
- [BrowseTo](Access.DoCmd.BrowseTo.md)
- [CancelEvent](Access.DoCmd.CancelEvent.md)
- [ClearMacroError](Access.DoCmd.ClearMacroError.md)
- [Close](Access.DoCmd.Close.md)
- [CloseDatabase](Access.DoCmd.CloseDatabase.md)
- [CopyDatabaseFile](Access.DoCmd.CopyDatabaseFile.md)
- [CopyObject](Access.DoCmd.CopyObject.md)
- [DeleteObject](Access.DoCmd.DeleteObject.md)
- [DoMenuItem](Access.DoCmd.DoMenuItem.md)
- [Echo](Access.DoCmd.Echo.md)
- [FindNext](Access.DoCmd.FindNext.md)
- [FindRecord](Access.DoCmd.FindRecord.md)
- [GoToControl](Access.DoCmd.GoToControl.md)
- [GoToPage](Access.DoCmd.GoToPage.md)
- [GoToRecord](Access.DoCmd.GoToRecord.md)
- [Hourglass](Access.DoCmd.Hourglass.md)
- [LockNavigationPane](Access.DoCmd.LockNavigationPane.md)
- [Maximize](Access.DoCmd.Maximize.md)
- [Minimize](Access.DoCmd.Minimize.md)
- [MoveSize](Access.DoCmd.MoveSize.md)
- [NavigateTo](Access.DoCmd.NavigateTo.md)
- [OpenDataAccessPage](Access.DoCmd.OpenDataAccessPage.md)
- [OpenDiagram](Access.DoCmd.OpenDiagram.md)
- [OpenForm](Access.DoCmd.OpenForm.md)
- [OpenFunction](Access.DoCmd.OpenFunction.md)
- [OpenModule](Access.DoCmd.OpenModule.md)
- [OpenQuery](Access.DoCmd.OpenQuery.md)
- [OpenReport](Access.DoCmd.OpenReport.md)
- [OpenStoredProcedure](Access.DoCmd.OpenStoredProcedure.md)
- [OpenTable](Access.DoCmd.OpenTable.md)
- [OpenView](Access.DoCmd.OpenView.md)
- [OutputTo](Access.DoCmd.OutputTo.md)
- [PrintOut](Access.DoCmd.PrintOut.md)
- [Quit](Access.DoCmd.Quit.md)
- [RefreshRecord](Access.DoCmd.RefreshRecord.md)
- [Rename](Access.DoCmd.Rename.md)
- [RepaintObject](Access.DoCmd.RepaintObject.md)
- [Requery](Access.DoCmd.Requery.md)
- [Restore](Access.DoCmd.Restore.md)
- [RunCommand](Access.DoCmd.RunCommand.md)
- [RunDataMacro](Access.DoCmd.RunDataMacro.md)
- [RunMacro](Access.DoCmd.RunMacro.md)
- [RunSavedImportExport](Access.DoCmd.RunSavedImportExport.md)
- [RunSQL](Access.DoCmd.RunSQL.md)
- [Save](Access.DoCmd.Save.md)
- [SearchForRecord](Access.DoCmd.SearchForRecord.md)
- [SelectObject](Access.DoCmd.SelectObject.md)
- [SendObject](Access.DoCmd.SendObject.md)
- [SetDisplayedCategories](Access.DoCmd.SetDisplayedCategories.md)
- [SetFilter](Access.DoCmd.SetFilter.md)
- [SetMenuItem](Access.DoCmd.SetMenuItem.md)
- [SetOrderBy](Access.DoCmd.SetOrderBy.md)
- [SetParameter](Access.DoCmd.SetParameter.md)
- [SetProperty](Access.DoCmd.SetProperty.md)
- [SetWarnings](Access.DoCmd.SetWarnings.md)
- [ShowAllRecords](Access.DoCmd.ShowAllRecords.md)
- [ShowToolbar](Access.DoCmd.ShowToolbar.md)
- [SingleStep](Access.DoCmd.SingleStep.md)
- [TransferDatabase](Access.DoCmd.TransferDatabase.md)
- [TransferSharePointList](Access.DoCmd.TransferSharePointList.md)
- [TransferSpreadsheet](Access.DoCmd.TransferSpreadsheet.md)
- [TransferSQLDatabase](Access.DoCmd.TransferSQLDatabase.md)
- [TransferText](Access.DoCmd.TransferText.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
