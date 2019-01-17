---
title: Application Object (Project)
ms.prod: project-server
api_name:
- Project.Application
ms.assetid: 8eb91712-7784-a102-38c0-19bb056c27e9
ms.date: 06/08/2017
localization_priority: Priority
---


# Application Object (Project)

Represents the entire Project application. The  **Application** object contains:


- Application-wide settings and options (many of the options in the  **Options** dialog box on the **Tools** menu, for example).
    
- Properties that return top-level objects, such as  **ActiveCell**, **ActiveProject**, and so forth.
    
- Methods that act on application-wide elements, such as views, selections, editing actions, and so forth.
    

## Using the Application Object

Use the  **[Application](./Project.Project.Application.md)** property to return an **Application** object in Project . The following example applies the **Windows** property to the **Application** object.


```vb
Application.Windows("Project1.mpp").Activate
```


## Using Project From Another Application: Late Binding

The following example creates the Microsoft Project  **Application** object at run time, creates a new project, adds a task, saves the project, and then closes the Project . For example, copy and paste the **CreateProject_Late** macro to the **ThisDocument** module in the Visual Basic Editor (VBE) of Word.


 **Note**  Because the application queries the  **MSProject.Application** type library only at run time, Microsoft IntelliSense is not available and performance is relatively poor with late binding. Scripting languages, such as JavaScript and VBScript, require late binding. VBScript supports only the generic **Object** and **Variant** data types. For better performance in VBA and other compiled languages, you should use early binding by setting a reference to the Project type library.


```vb
Sub CreateProject_Late() 
    Dim pjApp As Object 
    Set pjApp = CreateObject("MSProject.Application") 
    pjApp.Visible = True 
    pjApp.FileNew 
    pjApp.ActiveProject.Tasks.Add "Hang clocks" 
    pjApp.FileSaveAs "Clocks.mpp" 
    pjApp.FileClose 
    pjApp.Quit 
End Sub
```

If you do not set the  **Visible** property to **True**, the Project application operates in the background without being visible.


## Using Project From Another Application: Early Binding

Early binding has better performance because it loads the type library at design time. To use early binding, you must set a reference to the Project application from the application you are working in. For example, in the VBE for a Word document, click  **References** on the **Tools** menu, scroll through the **Available References** list, and then choose the **Microsoft Project 15.0 Object Library** checkbox.

The following example opens a project from another application such as Excel , adds a task, and then saves and closes the project. 




```vb
Sub ModifyProject_Early() 
    Dim pjApp As MSProject.Application 
    Set pjApp = New MSProject.Application 
    pjApp.Visible = True 
    pjApp.FileOpen "Clocks.mpp" 
    pjApp.ActiveProject.Tasks.Add "Wind clocks" 
    pjApp.FileSave 
    pjApp.FileClose 
    pjApp.Quit 
End Sub
```


## Remarks




 **Important**  For application-level events, register event handlers  _after_ you set `Application.Visible = True`.



If you instantiate Project from another application and register an application-level event before setting the  **Visible** property of the **Application** object to **True**, the properties and methods of child objects of **Application** do not work. For example, `Application.ActiveProject.Name` is not accessible.

Many of the properties and methods that return the most common user-interface objects, such as the active project—represented by the  **[ActiveProject](./Project.Application.ActiveProject.md)** property—can be used without the **Application** object qualifier. For example, instead of writing `Application.ActiveProject.Visible = True` you can write `ActiveProject.Visible = True`


## Events



|Name|
|:-----|
|[AfterCubeBuilt](./Project.Application.AfterCubeBuilt.md)|
|[ApplicationBeforeClose](./Project.Application.ApplicationBeforeClose.md)|
|[ConnectionStatusChanged](./Project.Application.ConnectionStatusChanged.md)|
|[IsFunctionalitySupported](./Project.Application.IsFunctionalitySupported.md)|
|[JobCompleted](./Project.Application.JobCompleted.md)|
|[JobStart](./Project.Application.JobStart.md)|
|[LoadWebPage](./Project.Application.LoadWebPage.md)|
|[LoadWebPane](./Project.Application.LoadWebPane.md)|
|[NewProject](./Project.Application.NewProject.md)|
|[OnUndoOrRedo](./Project.Application.OnUndoOrRedo.md)|
|[PaneActivate](./Project.Application.PaneActivate.md)|
|[ProjectAfterSave](./Project.Application.ProjectAfterSave.md)|
|[ProjectAssignmentNew](./Project.Application.ProjectAssignmentNew.md)|
|[ProjectBeforeAssignmentChange](./Project.Application.ProjectBeforeAssignmentChange.md)|
|[ProjectBeforeAssignmentChange2](./Project.Application.ProjectBeforeAssignmentChange2.md)|
|[ProjectBeforeAssignmentDelete](./Project.Application.ProjectBeforeAssignmentDelete.md)|
|[ProjectBeforeAssignmentDelete2](./Project.Application.ProjectBeforeAssignmentDelete2.md)|
|[ProjectBeforeAssignmentNew](./Project.Application.ProjectBeforeAssignmentNew.md)|
|[ProjectBeforeAssignmentNew2](./Project.Application.ProjectBeforeAssignmentNew2.md)|
|[ProjectBeforeClearBaseline](./Project.Application.ProjectBeforeClearBaseline.md)|
|[ProjectBeforeClose](./Project.Application.ProjectBeforeClose.md)|
|[ProjectBeforeClose2](./Project.Application.ProjectBeforeClose2.md)|
|[ProjectBeforePrint](./Project.Application.ProjectBeforePrint.md)|
|[ProjectBeforePrint2](./Project.Application.ProjectBeforePrint2.md)|
|[ProjectBeforePublish](./Project.Application.ProjectBeforePublish.md)|
|[ProjectBeforeResourceChange](./Project.Application.ProjectBeforeResourceChange.md)|
|[ProjectBeforeResourceChange2](./Project.Application.ProjectBeforeResourceChange2.md)|
|[ProjectBeforeResourceDelete](./Project.Application.ProjectBeforeResourceDelete.md)|
|[ProjectBeforeResourceDelete2](./Project.Application.ProjectBeforeResourceDelete2.md)|
|[ProjectBeforeResourceNew](./Project.Application.ProjectBeforeResourceNew.md)|
|[ProjectBeforeResourceNew2](./Project.Application.ProjectBeforeResourceNew2.md)|
|[ProjectBeforeSave](./Project.Application.ProjectBeforeSave.md)|
|[ProjectBeforeSave2](./Project.Application.ProjectBeforeSave2.md)|
|[ProjectBeforeSaveBaseline](./Project.Application.ProjectBeforeSaveBaseline.md)|
|[ProjectBeforeTaskChange](./Project.Application.ProjectBeforeTaskChange.md)|
|[ProjectBeforeTaskChange2](./Project.Application.ProjectBeforeTaskChange2.md)|
|[ProjectBeforeTaskDelete](./Project.Application.ProjectBeforeTaskDelete.md)|
|[ProjectBeforeTaskDelete2](./Project.Application.ProjectBeforeTaskDelete2.md)|
|[ProjectBeforeTaskNew](./Project.Application.ProjectBeforeTaskNew.md)|
|[ProjectBeforeTaskNew2](./Project.Application.ProjectBeforeTaskNew2.md)|
|[ProjectCalculate](./Project.Application.ProjectCalculate.md)|
|[ProjectResourceNew](./Project.Application.ProjectResourceNew.md)|
|[ProjectTaskNew](./Project.Application.ProjectTaskNew.md)|
|[SaveCompletedToServer](./Project.Application.SaveCompletedToServer.md)|
|[SaveStartingToServer](./Project.Application.SaveStartingToServer.md)|
|[SecondaryViewChange](./Project.Application.SecondaryViewChange.md)|
|[WindowActivate](./Project.Application.WindowActivate(even).md)|
|[WindowBeforeViewChange](./Project.Application.WindowBeforeViewChange.md)|
|[WindowDeactivate](./Project.Application.WindowDeactivate.md)|
|[WindowGoalAreaChange](./Project.Application.WindowGoalAreaChange.md)|
|[WindowSelectionChange](./Project.Application.WindowSelectionChange.md)|
|[WindowSidepaneDisplayChange](./Project.Application.WindowSidepaneDisplayChange.md)|
|[WindowSidepaneTaskChange](./Project.Application.WindowSidepaneTaskChange.md)|
|[WindowViewChange](./Project.Application.WindowViewChange.md)|
|[WorkpaneDisplayChange](./Project.Application.WorkpaneDisplayChange.md)|

## Methods



|Name|
|:-----|
|[About](./Project.Application.About.md)|
|[ActivateMicrosoftApp](./Project.Application.ActivateMicrosoftApp.md)|
|[AddNewColumn](./Project.Application.AddNewColumn.md)|
|[AddProgressLine](./Project.Application.AddProgressLine.md)|
|[AddResourcesFromProjectServer](./Project.Application.AddResourcesFromProjectServer.md)|
|[AddSiteColumn](./Project.application.addsitecolumn.md)|
|[AfterUnloadWebBrowserControl](./Project.Application.AfterUnloadWebBrowserControl.md)|
|[Alerts](./Project.Application.Alerts.md)|
|[AlignTableCellBottom](./Project.application.aligntablecellbottom.md)|
|[AlignTableCellTop](./Project.application.aligntablecelltop.md)|
|[AlignTableCellVerticalCenter](./Project.application.aligntablecellverticalcenter.md)|
|[AppExecute](./Project.Application.AppExecute.md)|
|[ApplyReport](./Project.application.applyreport.md)|
|[ApplyReportLayoutTemplate](./Project.application.applyreportlayouttemplate.md)|
|[AppMaximize](./Project.Application.AppMaximize.md)|
|[AppMinimize](./Project.Application.AppMinimize.md)|
|[AppMove](./Project.Application.AppMove.md)|
|[AppRestore](./Project.Application.AppRestore.md)|
|[AppSize](./Project.Application.AppSize.md)|
|[AutoCorrect](./Project.Application.AutoCorrect.md)|
|[AutoFilter](./Project.Application.AutoFilter.md)|
|[AutoSaveToGlobal](./Project.Application.AutoSaveToGlobal.md)|
|[BarBoxFormat](./Project.Application.BarBoxFormat.md)|
|[BarBoxStyles](./Project.Application.BarBoxStyles.md)|
|[BarRounding](./Project.Application.BarRounding.md)|
|[BaseCalendarCreate](./Project.Application.BaseCalendarCreate.md)|
|[BaseCalendarDelete](./Project.Application.BaseCalendarDelete.md)|
|[BaseCalendarEditDays](./Project.Application.BaseCalendarEditDays.md)|
|[BaseCalendarRename](./Project.Application.BaseCalendarRename.md)|
|[BaseCalendarReset](./Project.Application.BaseCalendarReset.md)|
|[BaseCalendars](./Project.Application.BaseCalendars.md)|
|[BaselineClear](./Project.Application.BaselineClear.md)|
|[BaselineSave](./Project.Application.BaselineSave.md)|
|[BoxAlign](./Project.Application.BoxAlign.md)|
|[BoxCellEdit](./Project.Application.BoxCellEdit.md)|
|[BoxCellEditEx](./Project.Application.BoxCellEditEx.md)|
|[BoxCellLayout](./Project.Application.BoxCellLayout.md)|
|[BoxDataTemplate](./Project.Application.BoxDataTemplate.md)|
|[BoxFormat](./Project.Application.BoxFormat.md)|
|[BoxFormatEx](./Project.Application.BoxFormatEx.md)|
|[BoxGetXPosition](./Project.Application.BoxGetXPosition.md)|
|[BoxGetYPosition](./Project.Application.BoxGetYPosition.md)|
|[BoxLayout](./Project.Application.BoxLayout.md)|
|[BoxLayoutEx](./Project.Application.BoxLayoutEx.md)|
|[BoxLinkLabelsShow](./Project.Application.BoxLinkLabelsShow.md)|
|[BoxLinks](./Project.Application.BoxLinks.md)|
|[BoxLinksEx](./Project.Application.BoxLinksEx.md)|
|[BoxLinkStyleToggle](./Project.Application.BoxLinkStyleToggle.md)|
|[BoxProgressMarksShow](./Project.Application.BoxProgressMarksShow.md)|
|[BoxSet](./Project.Application.BoxSet.md)|
|[BoxShowHideFields](./Project.Application.BoxShowHideFields.md)|
|[BoxStylesEdit](./Project.Application.BoxStylesEdit.md)|
|[BoxStylesEditEx](./Project.Application.BoxStylesEditEx.md)|
|[BoxZoom](./Project.Application.BoxZoom.md)|
|[CacheSettings](./Project.Application.CacheSettings.md)|
|[CacheStatus](./Project.Application.CacheStatus.md)|
|[CalculateAll](./Project.Application.CalculateAll.md)|
|[CalculateProject](./Project.Application.CalculateProject.md)|
|[CalendarBarStyles](./Project.Application.CalendarBarStyles.md)|
|[CalendarBarStylesEdit](./Project.Application.CalendarBarStylesEdit.md)|
|[CalendarBarStylesEditEx](./Project.Application.CalendarBarStylesEditEx.md)|
|[CalendarBestFitWeekHeight](./Project.Application.CalendarBestFitWeekHeight.md)|
|[CalendarDateBoxes](./Project.Application.CalendarDateBoxes.md)|
|[CalendarDateBoxesEx](./Project.Application.CalendarDateBoxesEx.md)|
|[CalendarDateShading](./Project.Application.CalendarDateShading.md)|
|[CalendarDateShadingEdit](./Project.Application.CalendarDateShadingEdit.md)|
|[CalendarDateShadingEditEx](./Project.Application.CalendarDateShadingEditEx.md)|
|[CalendarLayout](./Project.Application.CalendarLayout.md)|
|[CalendarShowBarSplits](./Project.Application.CalendarShowBarSplits.md)|
|[CalendarTaskList](./Project.Application.CalendarTaskList.md)|
|[CalendarTimescale](./Project.Application.CalendarTimescale.md)|
|[CalendarWeekHeadingsEx](./Project.Application.CalendarWeekHeadingsEx.md)|
|[ChangeColumnDataType](./Project.Application.ChangeColumnDataType.md)|
|[ChangeStatusDate](./Project.Application.ChangeStatusDate.md)|
|[ChangeWorkingTimeEx](./Project.Application.ChangeWorkingTimeEx.md)|
|[CheckField](./Project.Application.CheckField.md)|
|[CheckIn](./Project.Application.CheckIn.md)|
|[CheckOut](./Project.Application.CheckOut.md)|
|[CheckResourceErrors](./Project.Application.CheckResourceErrors.md)|
|[CheckTaskErrors](./Project.Application.CheckTaskErrors.md)|
|[CleanupCache](./Project.Application.CleanupCache.md)|
|[CleanupProjectFromCache](./Project.Application.CleanupProjectFromCache.md)|
|[ClearConstraint](./Project.Application.ClearConstraint.md)|
|[CloseComparison](./Project.Application.CloseComparison.md)|
|[CloseUndoTransaction](./Project.Application.CloseUndoTransaction.md)|
|[ColumnAlignment](./Project.Application.ColumnAlignment.md)|
|[ColumnBestFit](./Project.Application.ColumnBestFit.md)|
|[ColumnDelete](./Project.Application.ColumnDelete.md)|
|[ColumnEdit](./Project.Application.ColumnEdit.md)|
|[ColumnInsert](./Project.Application.ColumnInsert.md)|
|[ComAddInsDialog](./Project.Application.ComAddInsDialog.md)|
|[CommitmentsPane](./Project.Application.CommitmentsPane.md)|
|[CompareProjectsLegendToggle](./Project.Application.CompareProjectsLegendToggle.md)|
|[CompareProjectVersions](./Project.Application.CompareProjectVersions.md)|
|[ConsolidateProjects](./Project.Application.ConsolidateProjects.md)|
|[ConvertHangulToHanja](./Project.Application.ConvertHangulToHanja.md)|
|[CopyReport](./Project.application.copyreport.md)|
|[CreateComparisonReport](./Project.Application.CreateComparisonReport.md)|
|[CreateEnterpriseCalendar](./Project.Application.CreateEnterpriseCalendar.md)|
|[CreateProjectSite](./Project.application.createprojectsite.md)|
|[CustomFieldDelete](./Project.Application.CustomFieldDelete.md)|
|[CustomFieldGetFormula](./Project.Application.CustomFieldGetFormula.md)|
|[CustomFieldGetName](./Project.Application.CustomFieldGetName.md)|
|[CustomFieldIndicatorAdd](./Project.Application.CustomFieldIndicatorAdd.md)|
|[CustomFieldIndicatorDelete](./Project.Application.CustomFieldIndicatorDelete.md)|
|[CustomFieldIndicators](./Project.Application.CustomFieldIndicators.md)|
|[CustomFieldMappingDialog](./Project.Application.CustomFieldMappingDialog.md)|
|[CustomFieldPropertiesEx](./Project.Application.CustomFieldPropertiesEx.md)|
|[CustomFieldRename](./Project.Application.CustomFieldRename.md)|
|[CustomFieldSetFormula](./Project.Application.CustomFieldSetFormula.md)|
|[CustomFieldValueList](./Project.Application.CustomFieldValueList.md)|
|[CustomFieldValueListAdd](./Project.Application.CustomFieldValueListAdd.md)|
|[CustomFieldValueListDelete](./Project.Application.CustomFieldValueListDelete.md)|
|[CustomFieldValueListGetItem](./Project.Application.CustomFieldValueListGetItem.md)|
|[CustomForms](./Project.Application.CustomForms.md)|
|[CustomizeField](./Project.Application.CustomizeField.md)|
|[CustomizeIMEMode](./Project.Application.CustomizeIMEMode.md)|
|[CustomOutlineCodeEditEx](./Project.Application.CustomOutlineCodeEditEx.md)|
|[DateAdd](./Project.Application.DateAdd.md)|
|[DateDifference](./Project.Application.DateDifference.md)|
|[DateFormat](./Project.Application.DateFormat.md)|
|[DateSubtract](./Project.Application.DateSubtract.md)|
|[DDEExecute](./Project.Application.DDEExecute.md)|
|[DDEInitiate](./Project.Application.DDEInitiate.md)|
|[DDELinksUpdate](./Project.Application.DDELinksUpdate.md)|
|[DDEPasteLink](./Project.Application.DDEPasteLink.md)|
|[DDETerminate](./Project.Application.DDETerminate.md)|
|[DeleteFromDatabase](./Project.Application.DeleteFromDatabase.md)|
|[DependenciesPane](./Project.Application.DependenciesPane.md)|
|[DetailsPaneToggle](./Project.Application.DetailsPaneToggle.md)|
|[DetailStylesAdd](./Project.Application.DetailStylesAdd.md)|
|[DetailStylesFormat](./Project.Application.DetailStylesFormat.md)|
|[DetailStylesFormatEx](./Project.Application.DetailStylesFormatEx.md)|
|[DetailStylesProperties](./Project.Application.DetailStylesProperties.md)|
|[DetailStylesRemove](./Project.Application.DetailStylesRemove.md)|
|[DetailStylesRemoveAll](./Project.Application.DetailStylesRemoveAll.md)|
|[DetailStylesToggleItem](./Project.Application.DetailStylesToggleItem.md)|
|[DisplaySharedWorkspace](./Project.Application.DisplaySharedWorkspace.md)|
|[DistributeTableColumns](./Project.application.distributetablecolumns.md)|
|[DistributeTableRows](./Project.application.distributetablerows.md)|
|[DocClose](./Project.Application.DocClose.md)|
|[DocMaximize](./Project.Application.DocMaximize.md)|
|[DocMove](./Project.Application.DocMove.md)|
|[DocRestore](./Project.Application.DocRestore.md)|
|[DocSize](./Project.Application.DocSize.md)|
|[DocumentExport](./Project.Application.DocumentExport.md)|
|[DocumentLibraryVersionsDialog](./Project.Application.DocumentLibraryVersionsDialog.md)|
|[DrawingCreate](./Project.Application.DrawingCreate.md)|
|[DrawingCycleColor](./Project.Application.DrawingCycleColor.md)|
|[DrawingMove](./Project.Application.DrawingMove.md)|
|[DrawingProperties](./Project.Application.DrawingProperties.md)|
|[DrawingReshape](./Project.Application.DrawingReshape.md)|
|[DurationFormat](./Project.Application.DurationFormat.md)|
|[DurationValue](./Project.Application.DurationValue.md)|
|[EditClear](./Project.Application.EditClear.md)|
|[EditClearFormats](./Project.Application.EditClearFormats.md)|
|[EditClearHyperlink](./Project.Application.EditClearHyperlink.md)|
|[EditCopy](./Project.Application.EditCopy.md)|
|[EditCopyPicture](./Project.Application.EditCopyPicture.md)|
|[EditCut](./Project.Application.EditCut.md)|
|[EditDelete](./Project.Application.EditDelete.md)|
|[EditEnterpriseCalendar](./Project.Application.EditEnterpriseCalendar.md)|
|[EditGoTo](./Project.Application.EditGoTo.md)|
|[EditHyperlink](./Project.Application.EditHyperlink.md)|
|[EditInsert](./Project.Application.EditInsert.md)|
|[EditPaste](./Project.Application.EditPaste.md)|
|[EditPasteAsHyperlink](./Project.Application.EditPasteAsHyperlink.md)|
|[EditPasteSpecial](./Project.Application.EditPasteSpecial.md)|
|[EditRedo](./Project.Application.EditRedo.md)|
|[EditTPStyle](./Project.Application.EditTPStyle.md)|
|[EditUndo](./Project.Application.EditUndo.md)|
|[EnterpriseGlobalCheckOut](./Project.Application.EnterpriseGlobalCheckOut.md)|
|[EnterpriseMakeServerURLTrusted](./Project.Application.EnterpriseMakeServerURLTrusted.md)|
|[EnterpriseProjectDelete](./Project.Application.EnterpriseProjectDelete.md)|
|[EnterpriseProjectImportWizard](./Project.Application.EnterpriseProjectImportWizard.md)|
|[EnterpriseProjectProfiles](./Project.Application.EnterpriseProjectProfiles.md)|
|[EnterpriseResourceGet](./Project.Application.EnterpriseResourceGet.md)|
|[EnterpriseResourcesImportEx](./Project.Application.EnterpriseResourcesImportEx.md)|
|[EnterpriseResourcesOpen](./Project.Application.EnterpriseResourcesOpen.md)|
|[EnterpriseResSubstitutionWizard](./Project.Application.EnterpriseResSubstitutionWizard.md)|
|[EnterpriseTeamBuilder](./Project.Application.EnterpriseTeamBuilder.md)|
|[FieldConstantToFieldName](./Project.Application.FieldConstantToFieldName.md)|
|[FieldNameToFieldConstant](./Project.Application.FieldNameToFieldConstant.md)|
|[FileCloseAllEx](./Project.Application.FileCloseAllEx.md)|
|[FileCloseEx](./Project.Application.FileCloseEx.md)|
|[FileExit](./Project.Application.FileExit.md)|
|[FileLoadLast](./Project.Application.FileLoadLast.md)|
|[FileNew](./Project.Application.FileNew.md)|
|[FileOpenEx](./Project.Application.FileOpenEx.md)|
|[FileOpenOrCreate](./Project.application.fileopenorcreate.md)|
|[FileOpenUsingBackstage](./Project.application.fileopenusingbackstage.md)|
|[FilePageSetup](./Project.Application.FilePageSetup.md)|
|[FilePageSetupCalendar](./Project.Application.FilePageSetupCalendar.md)|
|[FilePageSetupCalendarText](./Project.Application.FilePageSetupCalendarText.md)|
|[FilePageSetupCalendarTextEx](./Project.Application.FilePageSetupCalendarTextEx.md)|
|[FilePageSetupFooter](./Project.Application.FilePageSetupFooter.md)|
|[FilePageSetupHeader](./Project.Application.FilePageSetupHeader.md)|
|[FilePageSetupLegend](./Project.Application.FilePageSetupLegend.md)|
|[FilePageSetupLegendEx](./Project.Application.FilePageSetupLegendEx.md)|
|[FilePageSetupMargins](./Project.Application.FilePageSetupMargins.md)|
|[FilePageSetupPage](./Project.Application.FilePageSetupPage.md)|
|[FilePageSetupView](./Project.Application.FilePageSetupView.md)|
|[FilePrint](./Project.Application.FilePrint.md)|
|[FilePrintPreview](./Project.Application.FilePrintPreview.md)|
|[FilePrintSetup](./Project.Application.FilePrintSetup.md)|
|[FileProperties](./Project.Application.FileProperties.md)|
|[FileSave](./Project.Application.FileSave.md)|
|[FileSaveAs](./Project.Application.FileSaveAs.md)|
|[FileSaveOffline](./Project.Application.FileSaveOffline.md)|
|[FileSaveWorkspace](./Project.Application.FileSaveWorkspace.md)|
|[FillAcross](./Project.Application.FillAcross.md)|
|[FillDown](./Project.Application.FillDown.md)|
|[FilterApply](./Project.Application.FilterApply.md)|
|[FilterClear](./Project.Application.FilterClear.md)|
|[FilterEdit](./Project.Application.FilterEdit.md)|
|[FilterNew](./Project.Application.FilterNew.md)|
|[Filters](./Project.Application.Filters.md)|
|[FilterShowSummaryRows](./Project.Application.FilterShowSummaryRows.md)|
|[Find](./Project.Application.Find.md)|
|[FindEx](./Project.Application.FindEx.md)|
|[FindFile](./Project.Application.FindFile.md)|
|[FindNext](./Project.Application.FindNext.md)|
|[FindPrevious](./Project.Application.FindPrevious.md)|
|[FollowHyperlink](./Project.Application.FollowHyperlink.md)|
|[Font32Ex](./Project.Application.Font32Ex.md)|
|[FontBold](./Project.Application.FontBold.md)|
|[FontEx](./Project.Application.FontEx.md)|
|[FontItalic](./Project.Application.FontItalic.md)|
|[FontStrikethrough](./Project.Application.FontStrikethrough.md)|
|[FontUnderLine](./Project.Application.FontUnderLine.md)|
|[Form](./Project.Application.Form.md)|
|[FormatCopy](./Project.Application.FormatCopy.md)|
|[FormatPainter](./Project.Application.FormatPainter.md)|
|[FormatPaste](./Project.Application.FormatPaste.md)|
|[FormViewShow](./Project.Application.FormViewShow.md)|
|[GanttBarEditEx](./Project.Application.GanttBarEditEx.md)|
|[GanttBarFormat](./Project.Application.GanttBarFormat.md)|
|[GanttBarFormatEx](./Project.Application.GanttBarFormatEx.md)|
|[GanttBarLinks](./Project.Application.GanttBarLinks.md)|
|[GanttBarSize](./Project.Application.GanttBarSize.md)|
|[GanttBarStyleBaseline](./Project.Application.GanttBarStyleBaseline.md)|
|[GanttBarStyleCritical](./Project.Application.GanttBarStyleCritical.md)|
|[GanttBarStyleDelete](./Project.Application.GanttBarStyleDelete.md)|
|[GanttBarStyleEdit](./Project.Application.GanttBarStyleEdit.md)|
|[GanttBarStyleLate](./Project.Application.GanttBarStyleLate.md)|
|[GanttBarStyleSlack](./Project.Application.GanttBarStyleSlack.md)|
|[GanttBarStyleSlippage](./Project.Application.GanttBarStyleSlippage.md)|
|[GanttBarTextDateFormat](./Project.Application.GanttBarTextDateFormat.md)|
|[GanttChartWizard](./Project.Application.GanttChartWizard.md)|
|[GanttRollup](./Project.Application.GanttRollup.md)|
|[GanttShowBarSplits](./Project.Application.GanttShowBarSplits.md)|
|[GanttShowDrawings](./Project.Application.GanttShowDrawings.md)|
|[GetCellInfo](./Project.Application.GetCellInfo.md)|
|[GetCurrentTheme](./Project.Application.GetCurrentTheme.md)|
|[GetProjectServerSettingsEx](./Project.Application.GetProjectServerSettingsEx.md)|
|[GetProjectServerVersion](./Project.Application.GetProjectServerVersion.md)|
|[GetRedoListCount](./Project.Application.GetRedoListCount.md)|
|[GetRedoListItem](./Project.Application.GetRedoListItem.md)|
|[GetThemedColor](./Project.Application.GetThemedColor.md)|
|[GetUndoListCount](./Project.Application.GetUndoListCount.md)|
|[GetUndoListItem](./Project.Application.GetUndoListItem.md)|
|[GoalAreaChange](./Project.Application.GoalAreaChange.md)|
|[GoalAreaHighlight](./Project.Application.GoalAreaHighlight.md)|
|[GoalAreaTaskHighlight](./Project.Application.GoalAreaTaskHighlight.md)|
|[GoToItemInVersions](./Project.Application.GoToItemInVersions.md)|
|[GotoNextOverAllocation](./Project.Application.GotoNextOverAllocation.md)|
|[GotoTaskDates](./Project.Application.GotoTaskDates.md)|
|[Gridlines](./Project.Application.Gridlines.md)|
|[GridlinesEdit](./Project.Application.GridlinesEdit.md)|
|[GridlinesEditEx](./Project.Application.GridlinesEditEx.md)|
|[GroupApply](./Project.Application.GroupApply.md)|
|[GroupBy](./Project.Application.GroupBy.md)|
|[GroupClear](./Project.Application.GroupClear.md)|
|[GroupMaintainHierarchy](./Project.Application.GroupMaintainHierarchy.md)|
|[GroupNew](./Project.Application.GroupNew.md)|
|[Groups](./Project.Application.Groups.md)|
|[HelpAbout](./Project.Application.HelpAbout.md)|
|[HelpAnswerWizard](./Project.Application.HelpAnswerWizard.md)|
|[HelpContents](./Project.Application.HelpContents.md)|
|[HelpLaunch](./Project.Application.HelpLaunch.md)|
|[HelpTechnicalSupport](./Project.Application.HelpTechnicalSupport.md)|
|[HighlightDrivenSuccessors](./Project.application.highlightdrivensuccessors.md)|
|[HighlightDrivingPredecessors](./Project.application.highlightdrivingpredecessors.md)|
|[HighlightPredecessors](./Project.application.highlightpredecessors.md)|
|[HighlightSuccessors](./Project.application.highlightsuccessors.md)|
|[ImportCommitment](./Project.Application.ImportCommitment.md)|
|[ImportOutlookTasks](./Project.Application.ImportOutlookTasks.md)|
|[InactivateTaskToggle](./Project.Application.InactivateTaskToggle.md)|
|[InformationDialog](./Project.Application.InformationDialog.md)|
|[InsertBlankRow](./Project.Application.InsertBlankRow.md)|
|[InsertHyperlink](./Project.Application.InsertHyperlink.md)|
|[InsertManualTask](./Project.Application.InsertManualTask.md)|
|[InsertMilestoneTask](./Project.Application.InsertMilestoneTask.md)|
|[InsertNotes](./Project.Application.InsertNotes.md)|
|[InsertResource](./Project.Application.InsertResource.md)|
|[InsertScheduledTask](./Project.Application.InsertScheduledTask.md)|
|[InsertSummaryTask](./Project.Application.InsertSummaryTask.md)|
|[InsertTask](./Project.Application.InsertTask.md)|
|[IsCommandEnabled](./Project.Application.IsCommandEnabled.md)|
|[IsOfficeTaskPaneVisible](./Project.Application.IsOfficeTaskPaneVisible.md)|
|[IsOffline](./Project.Application.IsOffline.md)|
|[IsReducedFunctionalityMode](./Project.Application.IsReducedFunctionalityMode.md)|
|[IsUndoingOrRedoing](./Project.Application.IsUndoingOrRedoing.md)|
|[IsURLTrusted](./Project.Application.IsURLTrusted.md)|
|[Layout](./Project.Application.Layout.md)|
|[LayoutNow](./Project.Application.LayoutNow.md)|
|[LayoutRelatedNow](./Project.Application.LayoutRelatedNow.md)|
|[LayoutSelectionNow](./Project.Application.LayoutSelectionNow.md)|
|[LevelingClear](./Project.Application.LevelingClear.md)|
|[LevelingOptions](./Project.Application.LevelingOptions.md)|
|[LevelingOptionsEx](./Project.Application.LevelingOptionsEx.md)|
|[LevelNow](./Project.Application.LevelNow.md)|
|[LevelSelected](./Project.Application.LevelSelected.md)|
|[LinksBetweenProjects](./Project.Application.LinksBetweenProjects.md)|
|[LinkTasks](./Project.Application.LinkTasks.md)|
|[LinkTasksEdit](./Project.Application.LinkTasksEdit.md)|
|[LinkToTaskList](./Project.application.linktotasklist.md)|
|[LoadWebBrowserControlEx](./Project.Application.LoadWebBrowserControlEx.md)|
|[LoadWebPaneControl](./Project.Application.LoadWebPaneControl.md)|
|[LocaleID](./Project.Application.LocaleID.md)|
|[LookUpTableAddEx](./Project.Application.LookUpTableAddEx.md)|
|[Macro](./Project.Application.Macro.md)|
|[MacroSecurity](./Project.Application.MacroSecurity.md)|
|[MacroShowCode](./Project.Application.MacroShowCode.md)|
|[MacroShowVba](./Project.Application.MacroShowVba.md)|
|[MailLogoff](./Project.Application.MailLogoff.md)|
|[MailLogon](./Project.Application.MailLogon.md)|
|[MailPostDocument](./Project.Application.MailPostDocument.md)|
|[MailRoutingSlip](./Project.Application.MailRoutingSlip.md)|
|[MailSend](./Project.Application.MailSend.md)|
|[MailSession](./Project.Application.MailSession.md)|
|[MailSystem](./Project.Application.MailSystem.md)|
|[MakeFieldEnterprise](./Project.Application.MakeFieldEnterprise.md)|
|[MakeLocalCalendarEnterprise](./Project.Application.MakeLocalCalendarEnterprise.md)|
|[ManageSiteColumns](./Project.Application.ManageSiteColumns.md)|
|[MapEdit](./Project.Application.MapEdit.md)|
|[Message](./Project.Application.Message.md)|
|[NewTasksStartOn](./Project.Application.NewTasksStartOn.md)|
|[ObjectChangeIcon](./Project.Application.ObjectChangeIcon.md)|
|[ObjectConvert](./Project.Application.ObjectConvert.md)|
|[ObjectInsert](./Project.Application.ObjectInsert.md)|
|[ObjectLinks](./Project.Application.ObjectLinks.md)|
|[ObjectVerb](./Project.Application.ObjectVerb.md)|
|[OfficeOnTheWeb](./Project.Application.OfficeOnTheWeb.md)|
|[OfficeTaskPaneHide](./Project.Application.OfficeTaskPaneHide.md)|
|[OpenBrowser](./Project.application.openbrowser.md)|
|[OpenFromSharePoint](./Project.Application.OpenFromSharePoint.md)|
|[OpenServerPage](./Project.Application.OpenServerPage.md)|
|[OpenUndoTransaction](./Project.Application.OpenUndoTransaction.md)|
|[OpenXML](./Project.Application.OpenXML.md)|
|[OptionsCalculation](./Project.Application.OptionsCalculation.md)|
|[OptionsCalendar](./Project.Application.OptionsCalendar.md)|
|[OptionsEditEx](./Project.Application.OptionsEditEx.md)|
|[OptionsGeneralEx](./Project.Application.OptionsGeneralEx.md)|
|[OptionsInterfaceEx](./Project.Application.OptionsInterfaceEx.md)|
|[OptionsSave](./Project.Application.OptionsSave.md)|
|[OptionsSchedule](./Project.Application.OptionsSchedule.md)|
|[OptionsSecurityEx](./Project.Application.OptionsSecurityEx.md)|
|[OptionsSecurityTab](./Project.Application.OptionsSecurityTab.md)|
|[OptionsSpelling](./Project.Application.OptionsSpelling.md)|
|[OptionsViewEx](./Project.Application.OptionsViewEx.md)|
|[Organizer](./Project.Application.Organizer.md)|
|[OrganizerDeleteItem](./Project.Application.OrganizerDeleteItem.md)|
|[OrganizerMoveItem](./Project.Application.OrganizerMoveItem.md)|
|[OrganizerRenameItem](./Project.Application.OrganizerRenameItem.md)|
|[OutlineHideSubTasks](./Project.Application.OutlineHideSubTasks.md)|
|[OutlineIndent](./Project.Application.OutlineIndent.md)|
|[OutlineOutdent](./Project.Application.OutlineOutdent.md)|
|[OutlineShowAllTasks](./Project.Application.OutlineShowAllTasks.md)|
|[OutlineShowSubTasks](./Project.Application.OutlineShowSubTasks.md)|
|[OutlineShowTasks](./Project.Application.OutlineShowTasks.md)|
|[OutlineSymbolsToggle](./Project.Application.OutlineSymbolsToggle.md)|
|[PageBreakRemove](./Project.Application.PageBreakRemove.md)|
|[PageBreakSet](./Project.Application.PageBreakSet.md)|
|[PageBreaksRemoveAll](./Project.Application.PageBreaksRemoveAll.md)|
|[PageBreaksShow](./Project.Application.PageBreaksShow.md)|
|[PaneClose](./Project.Application.PaneClose.md)|
|[PaneCreate](./Project.Application.PaneCreate.md)|
|[PaneNext](./Project.Application.PaneNext.md)|
|[PanZoomPanTo](./Project.Application.PanZoomPanTo.md)|
|[PanZoomZoomTo](./Project.Application.PanZoomZoomTo.md)|
|[PasteAsPicture](./Project.application.pasteaspicture.md)|
|[PasteDestFormatting](./Project.application.pastedestformatting.md)|
|[PasteSourceFormatting](./Project.application.pastesourceformatting.md)|
|[ProgressLines](./Project.Application.ProgressLines.md)|
|[ProjectCheckOut](./Project.application.projectcheckout.md)|
|[ProjectMove](./Project.Application.ProjectMove.md)|
|[ProjectStatistics](./Project.Application.ProjectStatistics.md)|
|[ProjectSummaryInfo](./Project.Application.ProjectSummaryInfo.md)|
|[Publish](./Project.Application.Publish.md)|
|[Quit](./Project.Application.Quit.md)|
|[ReassignSelectedAssns](./Project.Application.ReassignSelectedAssns.md)|
|[RecurringTaskInsert](./Project.Application.RecurringTaskInsert.md)|
|[Redo](./Project.Application.Redo.md)|
|[RegisterProject](./Project.Application.RegisterProject.md)|
|[ReminderSet](./Project.Application.ReminderSet.md)|
|[RemoveHighlight](./Project.application.removehighlight.md)|
|[RenameReport](./Project.application.renamereport.md)|
|[Replace](./Project.Application.Replace.md)|
|[ReplaceEx](./Project.Application.ReplaceEx.md)|
|[ReportPrint](./Project.Application.ReportPrint.md)|
|[ReportPrintPreview](./Project.Application.ReportPrintPreview.md)|
|[Reports](./Project.Application.Reports.md)|
|[ReportsDialog](./Project.application.reportsdialog.md)|
|[RequestProgressInformation](./Project.Application.RequestProgressInformation.md)|
|[RescheduleToNextAvailable](./Project.Application.RescheduleToNextAvailable.md)|
|[ResetTPStyle](./Project.Application.ResetTPStyle.md)|
|[ResourceActiveDirectory](./Project.Application.ResourceActiveDirectory.md)|
|[ResourceAddressBook](./Project.Application.ResourceAddressBook.md)|
|[ResourceAssignment](./Project.Application.ResourceAssignment.md)|
|[ResourceAssignmentDialog](./Project.Application.ResourceAssignmentDialog.md)|
|[ResourceCalendarEditDays](./Project.Application.ResourceCalendarEditDays.md)|
|[ResourceCalendarReset](./Project.Application.ResourceCalendarReset.md)|
|[ResourceCalendars](./Project.Application.ResourceCalendars.md)|
|[ResourceComparison](./Project.Application.ResourceComparison.md)|
|[ResourceDetails](./Project.Application.ResourceDetails.md)|
|[ResourceGraphBarStyles](./Project.Application.ResourceGraphBarStyles.md)|
|[ResourceGraphBarStylesEx](./Project.Application.ResourceGraphBarStylesEx.md)|
|[ResourceMappingDialog](./Project.Application.ResourceMappingDialog.md)|
|[ResourceSharing](./Project.Application.ResourceSharing.md)|
|[ResourceSharingPoolAction](./Project.Application.ResourceSharingPoolAction.md)|
|[ResourceSharingPoolRefresh](./Project.Application.ResourceSharingPoolRefresh.md)|
|[ResourceSharingPoolUpdate](./Project.Application.ResourceSharingPoolUpdate.md)|
|[ResourceWindowsAccount](./Project.Application.ResourceWindowsAccount.md)|
|[RestoreSheetSelection](./Project.Application.RestoreSheetSelection.md)|
|[RowClear](./Project.Application.RowClear.md)|
|[RowDelete](./Project.Application.RowDelete.md)|
|[RowInsert](./Project.Application.RowInsert.md)|
|[Run](./Project.Application.Run.md)|
|[SaveForSharing](./Project.Application.SaveForSharing.md)|
|[SaveSheetSelection](./Project.Application.SaveSheetSelection.md)|
|[SegmentBorderColor](./Project.Application.SegmentBorderColor.md)|
|[SegmentFillColor](./Project.Application.SegmentFillColor.md)|
|[SelectAll](./Project.Application.SelectAll.md)|
|[SelectBeginning](./Project.Application.SelectBeginning.md)|
|[SelectCell](./Project.Application.SelectCell.md)|
|[SelectCellDown](./Project.Application.SelectCellDown.md)|
|[SelectCellLeft](./Project.Application.SelectCellLeft.md)|
|[SelectCellRight](./Project.Application.SelectCellRight.md)|
|[SelectCellUp](./Project.Application.SelectCellUp.md)|
|[SelectColumn](./Project.Application.SelectColumn.md)|
|[SelectEnd](./Project.Application.SelectEnd.md)|
|[SelectionExtend](./Project.Application.SelectionExtend.md)|
|[SelectRange](./Project.Application.SelectRange.md)|
|[SelectResourceCell](./Project.Application.SelectResourceCell.md)|
|[SelectResourceColumn](./Project.Application.SelectResourceColumn.md)|
|[SelectResourceField](./Project.Application.SelectResourceField.md)|
|[SelectRow](./Project.Application.SelectRow.md)|
|[SelectRowEnd](./Project.Application.SelectRowEnd.md)|
|[SelectRowStart](./Project.Application.SelectRowStart.md)|
|[SelectSheet](./Project.Application.SelectSheet.md)|
|[SelectTable](./Project.application.selecttable.md)|
|[SelectTaskAssns](./Project.Application.SelectTaskAssns.md)|
|[SelectTaskCell](./Project.Application.SelectTaskCell.md)|
|[SelectTaskColumn](./Project.Application.SelectTaskColumn.md)|
|[SelectTaskField](./Project.Application.SelectTaskField.md)|
|[SelectTimescaleRange](./Project.Application.SelectTimescaleRange.md)|
|[SelectToEnd](./Project.Application.SelectToEnd.md)|
|[SelectTPLineHeight](./Project.Application.SelectTPLineHeight.md)|
|[SelectTPTask](./Project.Application.SelectTPTask.md)|
|[ServiceOptionsDialog](./Project.Application.ServiceOptionsDialog.md)|
|[SetActiveCell](./Project.Application.SetActiveCell.md)|
|[SetAutoFilter](./Project.Application.SetAutoFilter.md)|
|[SetField](./Project.Application.SetField.md)|
|[SetLTRTable](./Project.application.setltrtable.md)|
|[SetMatchingField](./Project.Application.SetMatchingField.md)|
|[SetResourceField](./Project.Application.SetResourceField.md)|
|[SetResourceFieldByID](./Project.Application.SetResourceFieldByID.md)|
|[SetRowHeight](./Project.Application.SetRowHeight.md)|
|[SetRTLTable](./Project.application.setrtltable.md)|
|[SetShowTaskSuggestions](./Project.Application.SetShowTaskSuggestions.md)|
|[SetShowTaskWarnings](./Project.Application.SetShowTaskWarnings.md)|
|[SetSidepaneStateButton](./Project.Application.SetSidepaneStateButton.md)|
|[SetSplitBar](./Project.Application.SetSplitBar.md)|
|[SetTaskField](./Project.Application.SetTaskField.md)|
|[SetTaskFieldByID](./Project.Application.SetTaskFieldByID.md)|
|[SetTaskMode](./Project.Application.SetTaskMode.md)|
|[SetTitleRowHeight](./Project.Application.SetTitleRowHeight.md)|
|[SetTPField](./Project.Application.SetTPField.md)|
|[ShareProjectOnline](./Project.Application.ShareProjectOnline.md)|
|[ShowAddNewColumn](./Project.Application.ShowAddNewColumn.md)|
|[ShowIgnoredTaskWarnings](./Project.Application.ShowIgnoredTaskWarnings.md)|
|[ShowOSFTaskPane](./Project.application.showosftaskpane.md)|
|[ShowReportDataPane](./Project.application.showreportdatapane.md)|
|[SidepaneTaskChange](./Project.Application.SidepaneTaskChange.md)|
|[SidepaneToggle](./Project.Application.SidepaneToggle.md)|
|[Sort](./Project.Application.Sort.md)|
|[SpellCheckField](./Project.Application.SpellCheckField.md)|
|[SpellingCheck](./Project.Application.SpellingCheck.md)|
|[SplitTask](./Project.Application.SplitTask.md)|
|[StopWebBrowserControlNavigation](./Project.Application.StopWebBrowserControlNavigation.md)|
|[SummaryResourceAssignmentsRefresh](./Project.Application.SummaryResourceAssignmentsRefresh.md)|
|[SummaryTasksShow](./Project.Application.SummaryTasksShow.md)|
|[SynchronizeWithSite](./Project.Application.SynchronizeWithSite.md)|
|[Table](./Project.application.table.md)|
|[TableApply](./Project.Application.TableApply.md)|
|[TableCopy](./Project.Application.TableCopy.md)|
|[TableEdit](./Project.Application.TableEdit.md)|
|[TableEditEx](./Project.Application.TableEditEx.md)|
|[TableReset](./Project.Application.TableReset.md)|
|[Tables](./Project.Application.Tables.md)|
|[TaskComparison](./Project.Application.TaskComparison.md)|
|[TaskDeliverableCreate](./Project.Application.TaskDeliverableCreate.md)|
|[TaskDeliverableSync](./Project.Application.TaskDeliverableSync.md)|
|[TaskDependencySync](./Project.Application.TaskDependencySync.md)|
|[TaskDrivers](./Project.Application.TaskDrivers.md)|
|[TaskInspector](./Project.Application.TaskInspector.md)|
|[TaskMove](./Project.Application.TaskMove.md)|
|[TaskMoveToStatusDate](./Project.Application.TaskMoveToStatusDate.md)|
|[TaskOnTimeline](./Project.Application.TaskOnTimeline.md)|
|[TaskRespectLinks](./Project.Application.TaskRespectLinks.md)|
|[TextStyles32Ex](./Project.Application.TextStyles32Ex.md)|
|[TextStylesEx](./Project.Application.TextStylesEx.md)|
|[TimelineExport](./Project.Application.TimelineExport.md)|
|[TimelineFormat](./Project.Application.TimelineFormat.md)|
|[TimelineGotoSelectedTask](./Project.Application.TimelineGotoSelectedTask.md)|
|[TimelineInsertTask](./Project.Application.TimelineInsertTask.md)|
|[TimelineShowHide](./Project.Application.TimelineShowHide.md)|
|[TimelineTextOnBar](./Project.Application.TimelineTextOnBar.md)|
|[TimelineViewToggle](./Project.Application.TimelineViewToggle.md)|
|[Timescale](./Project.Application.Timescale.md)|
|[TimescaleEdit](./Project.Application.TimescaleEdit.md)|
|[TimescaleNonWorking](./Project.Application.TimescaleNonWorking.md)|
|[TimescaleNonWorkingEx](./Project.Application.TimescaleNonWorkingEx.md)|
|[ToggleAssignments](./Project.Application.ToggleAssignments.md)|
|[ToggleChangeHighlighting](./Project.Application.ToggleChangeHighlighting.md)|
|[TogglePreventResOveralloc](./Project.Application.TogglePreventResOveralloc.md)|
|[ToggleResourceDetails](./Project.Application.ToggleResourceDetails.md)|
|[ToggleTaskDetails](./Project.Application.ToggleTaskDetails.md)|
|[ToggleTPAutoExpand](./Project.Application.ToggleTPAutoExpand.md)|
|[ToggleTPResourceExpand](./Project.Application.ToggleTPResourceExpand.md)|
|[ToggleTPUnassigned](./Project.Application.ToggleTPUnassigned.md)|
|[ToggleTPUnscheduled](./Project.Application.ToggleTPUnscheduled.md)|
|[Undo](./Project.Application.Undo.md)|
|[UndoClear](./Project.Application.UndoClear.md)|
|[UnlinkTasks](./Project.Application.UnlinkTasks.md)|
|[UnloadWebBrowserControl](./Project.Application.UnloadWebBrowserControl.md)|
|[UpdateFromProjectServer](./Project.Application.UpdateFromProjectServer.md)|
|[UpdateProject](./Project.Application.UpdateProject.md)|
|[UpdateTasks](./Project.Application.UpdateTasks.md)|
|[UsageViewEntryEx](./Project.Application.UsageViewEntryEx.md)|
|[ViewApply](./Project.Application.ViewApply.md)|
|[ViewApplyEx](./Project.Application.ViewApplyEx.md)|
|[ViewBar](./Project.Application.ViewBar.md)|
|[ViewCopy](./Project.Application.ViewCopy.md)|
|[ViewEditCombination](./Project.Application.ViewEditCombination.md)|
|[ViewEditSingle](./Project.Application.ViewEditSingle.md)|
|[ViewReset](./Project.Application.ViewReset.md)|
|[Views](./Project.Application.Views.md)|
|[ViewsEx](./Project.Application.ViewsEx.md)|
|[ViewShowCost](./Project.Application.ViewShowCost.md)|
|[ViewShowCumulativeCost](./Project.Application.ViewShowCumulativeCost.md)|
|[ViewShowCumulativeWork](./Project.Application.ViewShowCumulativeWork.md)|
|[ViewShowNotes](./Project.Application.ViewShowNotes.md)|
|[ViewShowObjects](./Project.Application.ViewShowObjects.md)|
|[ViewShowOverallocation](./Project.Application.ViewShowOverallocation.md)|
|[ViewShowPeakUnits](./Project.Application.ViewShowPeakUnits.md)|
|[ViewShowPercentAllocation](./Project.Application.ViewShowPercentAllocation.md)|
|[ViewShowPredecessorsSuccessors](./Project.Application.ViewShowPredecessorsSuccessors.md)|
|[ViewShowRemainingAvailability](./Project.Application.ViewShowRemainingAvailability.md)|
|[ViewShowResourcesPredecessors](./Project.Application.ViewShowResourcesPredecessors.md)|
|[ViewShowResourcesSuccessors](./Project.Application.ViewShowResourcesSuccessors.md)|
|[ViewShowSchedule](./Project.Application.ViewShowSchedule.md)|
|[ViewShowUnitAvailability](./Project.Application.ViewShowUnitAvailability.md)|
|[ViewShowWork](./Project.Application.ViewShowWork.md)|
|[ViewShowWorkAvailability](./Project.Application.ViewShowWorkAvailability.md)|
|[VisualReports](./Project.Application.VisualReports.md)|
|[VisualReportsEdit](./Project.Application.VisualReportsEdit.md)|
|[VisualReportsNewTemplate](./Project.Application.VisualReportsNewTemplate.md)|
|[VisualReportsSaveCube](./Project.Application.VisualReportsSaveCube.md)|
|[VisualReportsSaveDatabase](./Project.Application.VisualReportsSaveDatabase.md)|
|[VisualReportsView](./Project.Application.VisualReportsView.md)|
|[WBSCodeMaskEdit](./Project.Application.WBSCodeMaskEdit.md)|
|[WBSCodeRenumber](./Project.Application.WBSCodeRenumber.md)|
|[WebAddToFavorites](./Project.Application.WebAddToFavorites.md)|
|[WebCopyHyperlink](./Project.Application.WebCopyHyperlink.md)|
|[WebGoBack](./Project.Application.WebGoBack.md)|
|[WebGoForward](./Project.Application.WebGoForward.md)|
|[WebHideToolbars](./Project.Application.WebHideToolbars.md)|
|[WebOpenFavorites](./Project.Application.WebOpenFavorites.md)|
|[WebOpenHyperlink](./Project.Application.WebOpenHyperlink.md)|
|[WebOpenSearchPage](./Project.Application.WebOpenSearchPage.md)|
|[WebOpenStartPage](./Project.Application.WebOpenStartPage.md)|
|[WebRefresh](./Project.Application.WebRefresh.md)|
|[WebSetSearchPage](./Project.Application.WebSetSearchPage.md)|
|[WebSetStartPage](./Project.Application.WebSetStartPage.md)|
|[WebStopLoading](./Project.Application.WebStopLoading.md)|
|[WebToolbar](./Project.Application.WebToolbar.md)|
|[WindowActivate](./Project.Application.WindowActivate(method).md)|
|[WindowArrangeAll](./Project.Application.WindowArrangeAll.md)|
|[WindowHide](./Project.Application.WindowHide.md)|
|[WindowMoreWindows](./Project.Application.WindowMoreWindows.md)|
|[WindowNewWindow](./Project.Application.WindowNewWindow.md)|
|[WindowNext](./Project.Application.WindowNext.md)|
|[WindowPrev](./Project.Application.WindowPrev.md)|
|[WindowSplit](./Project.Application.WindowSplit.md)|
|[WindowUnhide](./Project.Application.WindowUnhide.md)|
|[WorkOffline](./Project.Application.WorkOffline.md)|
|[WrapText](./Project.Application.WrapText.md)|
|[Zoom](./Project.Application.Zoom.md)|
|[ZoomCalendar](./Project.Application.ZoomCalendar.md)|
|[ZoomIn](./Project.Application.ZoomIn.md)|
|[ZoomOut](./Project.Application.ZoomOut.md)|
|[ZoomReport](./Project.application.zoomreport.md)|
|[ZoomTimescale](./Project.Application.ZoomTimescale.md)|
|[AddEngagement](./Project.application.addengagement.md)|
|[EngagementInfo](./Project.application.engagementinfo.md)|
|[GetDpiScaleFactor](./Project.application.getdpiscalefactor.md)|
|[InsertTimelineBar](./Project.application.addtimelinebar.md)|
|[Inspector](./Project.application.inspector.md)|
|[LocaleName](./Project.application.localename.md)|
|[ProjectSummaryInfoEx](./Project.application.projectsummaryinfoex.md)|
|[RefreshEngagementsForProject](./Project.application.refreshengagementsforproject.md)|
|[RemoveTimelineBar](./Project.application.removetimelinebar.md)|
|[SubmitAllEngagementsForProject](./Project.application.submitallengagementsforproject.md)|
|[SubmitSelectedEngagementsForProject](./Project.application.submitselectedengagementsforproject.md)|
|[TaskOnTimelineEx](./Project.application.taskontimelineex.md)|
|[TimelineBarDateRange](./Project.application.timelinebardaterange.md)|
|[UpdateEngagementsForProject](./Project.application.updateengagementsforproject.md)|

## Properties



|Name|
|:-----|
|[ActiveCell](./Project.Application.ActiveCell.md)|
|[ActiveProject](./Project.Application.ActiveProject.md)|
|[ActiveSelection](./Project.Application.ActiveSelection.md)|
|[ActiveWindow](./Project.Application.ActiveWindow.md)|
|[AMText](./Project.Application.AMText.md)|
|[Application](./Project.Application.Application.md)|
|[AskToUpdateLinks](./Project.Application.AskToUpdateLinks.md)|
|[Assistance](./Project.application.assistance.md)|
|[AutoClearLeveling](./Project.Application.AutoClearLeveling.md)|
|[AutoLevel](./Project.Application.AutoLevel.md)|
|[AutomaticallyFillPhoneticFields](./Project.Application.AutomaticallyFillPhoneticFields.md)|
|[AutomationSecurity](./Project.Application.AutomationSecurity.md)|
|[Build](./Project.Application.Build.md)|
|[Calculation](./Project.Application.Calculation.md)|
|[Caption](./Project.Application.Caption.md)|
|[CellDragAndDrop](./Project.Application.CellDragAndDrop.md)|
|[COMAddIns](./Project.Application.COMAddIns.md)|
|[CommandBars](./Project.Application.CommandBars.md)|
|[CompareProjectsCurrentVersionName](./Project.Application.CompareProjectsCurrentVersionName.md)|
|[CompareProjectsPreviousVersionName](./Project.Application.CompareProjectsPreviousVersionName.md)|
|[DateOrder](./Project.Application.DateOrder.md)|
|[DateSeparator](./Project.Application.DateSeparator.md)|
|[DayLeadingZero](./Project.Application.DayLeadingZero.md)|
|[DecimalSeparator](./Project.Application.DecimalSeparator.md)|
|[DefaultAutoFilter](./Project.Application.DefaultAutoFilter.md)|
|[DefaultDateFormat](./Project.Application.DefaultDateFormat.md)|
|[DefaultView](./Project.Application.DefaultView.md)|
|[DisplayAlerts](./Project.Application.DisplayAlerts.md)|
|[DisplayEntryBar](./Project.Application.DisplayEntryBar.md)|
|[DisplayOLEIndicator](./Project.Application.DisplayOLEIndicator.md)|
|[DisplayPlanningWizard](./Project.Application.DisplayPlanningWizard.md)|
|[DisplayProjectGuide](./Project.Application.DisplayProjectGuide.md)|
|[DisplayRecentFiles](./Project.Application.DisplayRecentFiles.md)|
|[DisplayScheduleMessages](./Project.Application.DisplayScheduleMessages.md)|
|[DisplayScrollBars](./Project.Application.DisplayScrollBars.md)|
|[DisplayStatusBar](./Project.Application.DisplayStatusBar.md)|
|[DisplayViewBar](./Project.Application.DisplayViewBar.md)|
|[DisplayWindowsInTaskbar](./Project.Application.DisplayWindowsInTaskbar.md)|
|[DisplayWizardErrors](./Project.Application.DisplayWizardErrors.md)|
|[DisplayWizardScheduling](./Project.Application.DisplayWizardScheduling.md)|
|[DisplayWizardUsage](./Project.Application.DisplayWizardUsage.md)|
|[Edition](./Project.Application.Edition.md)|
|[EnableCancelKey](./Project.Application.EnableCancelKey.md)|
|[EnableChangeHighlighting](./Project.Application.EnableChangeHighlighting.md)|
|[EnterpriseAllowLocalBaseCalendars](./Project.Application.EnterpriseAllowLocalBaseCalendars.md)|
|[EnterpriseListSeparator](./Project.Application.EnterpriseListSeparator.md)|
|[EnterpriseProtectActuals](./Project.Application.EnterpriseProtectActuals.md)|
|[FileBuildID](./Project.Application.FileBuildID.md)|
|[FileFormatID](./Project.Application.FileFormatID.md)|
|[GetCacheStatusForProject](./Project.application.getcachestatusforproject.md)|
|[GlobalBaseCalendars](./Project.Application.GlobalBaseCalendars.md)|
|[GlobalOutlineCodes](./Project.Application.GlobalOutlineCodes.md)|
|[GlobalReports](./Project.application.globalreports.md)|
|[GlobalResourceFilters](./Project.Application.GlobalResourceFilters.md)|
|[GlobalResourceTables](./Project.Application.GlobalResourceTables.md)|
|[GlobalTaskFilters](./Project.Application.GlobalTaskFilters.md)|
|[GlobalTaskTables](./Project.Application.GlobalTaskTables.md)|
|[GlobalViews](./Project.Application.GlobalViews.md)|
|[GlobalViewsCombination](./Project.Application.GlobalViewsCombination.md)|
|[GlobalViewsSingle](./Project.Application.GlobalViewsSingle.md)|
|[Height](./Project.Application.Height.md)|
|[IsCheckedOut](./Project.application.ischeckedout.md)|
|[Left](./Project.Application.Left.md)|
|[LevelFreeformTasks](./Project.Application.LevelFreeformTasks.md)|
|[LevelIndividualAssignments](./Project.Application.LevelIndividualAssignments.md)|
|[LevelingCanSplit](./Project.Application.LevelingCanSplit.md)|
|[LevelOrder](./Project.Application.LevelOrder.md)|
|[LevelPeriodBasis](./Project.Application.LevelPeriodBasis.md)|
|[LevelProposedBookings](./Project.Application.LevelProposedBookings.md)|
|[LevelWithinSlack](./Project.Application.LevelWithinSlack.md)|
|[ListSeparator](./Project.Application.ListSeparator.md)|
|[LoadLastFile](./Project.Application.LoadLastFile.md)|
|[MonthLeadingZero](./Project.Application.MonthLeadingZero.md)|
|[MoveAfterReturn](./Project.Application.MoveAfterReturn.md)|
|[Name](./Project.Application.Name.md)|
|[NewTasksEstimated](./Project.Application.NewTasksEstimated.md)|
|[OperatingSystem](./Project.Application.OperatingSystem.md)|
|[PanZoomFinish](./Project.Application.PanZoomFinish.md)|
|[PanZoomStart](./Project.Application.PanZoomStart.md)|
|[Parent](./Project.Application.Parent.md)|
|[Path](./Project.Application.Path.md)|
|[PathSeparator](./Project.Application.PathSeparator.md)|
|[PMText](./Project.Application.PMText.md)|
|[Profiles](./Project.Application.Profiles.md)|
|[Projects](./Project.Application.Projects.md)|
|[PromptForSummaryInfo](./Project.Application.PromptForSummaryInfo.md)|
|[RecentFilesMaximum](./Project.Application.RecentFilesMaximum.md)|
|[ScreenUpdating](./Project.Application.ScreenUpdating.md)|
|[ShowAssignmentUnitsAs](./Project.Application.ShowAssignmentUnitsAs.md)|
|[ShowEstimatedDuration](./Project.Application.ShowEstimatedDuration.md)|
|[ShowWelcome](./Project.Application.ShowWelcome.md)|
|[StartWeekOn](./Project.Application.StartWeekOn.md)|
|[StartYearIn](./Project.Application.StartYearIn.md)|
|[StatusBar](./Project.Application.StatusBar.md)|
|[SupportsMultipleDocuments](./Project.Application.SupportsMultipleDocuments.md)|
|[SupportsMultipleWindows](./Project.Application.SupportsMultipleWindows.md)|
|[ThousandSeparator](./Project.Application.ThousandSeparator.md)|
|[TimeLeadingZero](./Project.Application.TimeLeadingZero.md)|
|[TimescaleFinish](./Project.Application.TimescaleFinish.md)|
|[TimescaleStart](./Project.Application.TimescaleStart.md)|
|[TimeSeparator](./Project.Application.TimeSeparator.md)|
|[Top](./Project.Application.Top.md)|
|[TrustProjectServerAndWSSPages](./Project.Application.TrustProjectServerAndWSSPages.md)|
|[TwelveHourTimeFormat](./Project.Application.TwelveHourTimeFormat.md)|
|[UndoLevels](./Project.Application.UndoLevels.md)|
|[UsableHeight](./Project.Application.UsableHeight.md)|
|[UsableWidth](./Project.Application.UsableWidth.md)|
|[Use3DLook](./Project.Application.Use3DLook.md)|
|[UseOMIDs](./Project.Application.UseOMIDs.md)|
|[UserControl](./Project.Application.UserControl.md)|
|[UserName](./Project.Application.UserName.md)|
|[VBE](./Project.Application.VBE.md)|
|[Version](./Project.Application.Version.md)|
|[Visible](./Project.Application.Visible.md)|
|[VisualReportsAdditionalTemplatePath](./Project.Application.VisualReportsAdditionalTemplatePath.md)|
|[VisualReportTemplateList](./Project.Application.VisualReportTemplateList.md)|
|[Width](./Project.Application.Width.md)|
|[Windows](./Project.Application.Windows.md)|
|[Windows2](./Project.Application.Windows2.md)|
|[WindowState](./Project.Application.WindowState.md)|

