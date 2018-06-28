---
title: Application Object (Access)
keywords: vbaac10.chm12627
f1_keywords:
- vbaac10.chm12627
ms.prod: access
api_name:
- Access.Application
ms.assetid: aefb0713-97e6-e2c7-e530-8fd2e1316a55
ms.date: 06/08/2017
---


# Application Object (Access)

The  **Application** object refers to the active Microsoft Access application.


## Remarks

The  **Application** object contains all Access objects and collections.

You can use the  **Application** object to apply methods or property settings to the entire Access application. For example, you can use the **[SetOption](../../../api/Access.Application.SetOption.md)** method of the **Application** object to set database options from Visual Basic. The following example shows how you can set the **Display Status Bar** check box on the **Current Database** tab of the **Access Options** dialog box.




```
Application.SetOption "Show Status Bar", True
```

Access is a COM component that supports Automation, formerly called OLE Automation. You can manipulate Access objects from another application that also supports Automation. To do this, you use the  **Application** object.

For example, Microsoft Visual Basic is a COM component. You can open anAccess database from Visual Basic and work with its objects. From Visual Basic, first create a reference to the Access object library. Then create a new instance of the  **Application** class and point an object variable to it, as in the following example:




```
Dim appAccess As New Access.Application
```

From applications that do not support the  **New** keyword, you can create a new instance of the **Application** class by using the **CreateObject** function:




```
Dim appAccess As Object 
Set appAccess = CreateObject("Access.Application")
```

After you create a new instance of the  **Application** class, you can open a database or create a new database, by using either the **[OpenCurrentDatabase](../../../api/Access.Application.OpenCurrentDatabase.md)** method or the **[NewCurrentDatabase](../../../api/Access.Application.NewCurrentDatabase.md)** method. You can then set the properties of the **Application** object and call its methods. When you return a reference to the **CommandBars** object by using the **CommandBars** property of the **Application** object, you can access all Microsoft Office command bar objects and collections by using this reference.

You can also manipulate other Access objects through the  **Application** object. For example, by using the **[OpenForm](../../../api/Access.DoCmd.OpenForm.md)** method of the Access **[DoCmd](docmd-object-access.md)** object, you can open an Access form from Microsoft Office Excel:




```
appAccess.DoCmd.OpenForm "Orders"
```

For more information about creating a reference and controlling objects by using Automation, see the documentation for the application that is acting as the COM component.


## Methods



|**Name**|
|:-----|
|[AccessError](../../../api/Access.Application.AccessError.md)|
|[AddToFavorites](../../../api/Access.Application.AddToFavorites.md)|
|[BuildCriteria](../../../api/Access.Application.BuildCriteria.md)|
|[CloseCurrentDatabase](../../../api/Access.Application.CloseCurrentDatabase.md)|
|[CodeDb](../../../api/Access.Application.CodeDb.md)|
|[ColumnHistory](../../../api/Access.Application.ColumnHistory.md)|
|[CompactRepair](../../../api/Access.Application.CompactRepair.md)|
|[ConvertAccessProject](../../../api/Access.Application.ConvertAccessProject.md)|
|[CreateAccessProject](../../../api/Access.Application.CreateAccessProject.md)|
|[CreateAdditionalData](../../../api/Access.Application.CreateAdditionalData.md)|
|[CreateControl](../../../api/Access.Application.CreateControl.md)|
|[CreateForm](../../../api/Access.Application.CreateForm.md)|
|[CreateGroupLevel](../../../api/Access.Application.CreateGroupLevel.md)|
|[CreateReport](../../../api/Access.Application.CreateReport.md)|
|[CreateReportControl](../../../api/Access.Application.CreateReportControl.md)|
|[CurrentDb](../../../api/Access.Application.CurrentDb.md)|
|[CurrentUser](../../../api/Access.Application.CurrentUser.md)|
|[CurrentWebUser](../../../api/Access.Application.CurrentWebUser.md)|
|[CurrentWebUserGroups](../../../api/Access.Application.CurrentWebUserGroups.md)|
|[DAvg](../../../api/Access.application.davg.md)|
|[DCount](../../../api/Access.Application.DCount.md)|
|[DDEExecute](../../../api/Access.Application.DDEExecute.md)|
|[DDEInitiate](../../../api/Access.Application.DDEInitiate.md)|
|[DDEPoke](../../../api/Access.Application.DDEPoke.md)|
|[DDERequest](../../../api/Access.Application.DDERequest.md)|
|[DDETerminate](../../../api/Access.Application.DDETerminate.md)|
|[DDETerminateAll](../../../api/Access.Application.DDETerminateAll.md)|
|[DefaultWorkspaceClone](../../../api/Access.Application.DefaultWorkspaceClone.md)|
|[DeleteControl](../../../api/Access.Application.DeleteControl.md)|
|[DeleteReportControl](../../../api/Access.Application.DeleteReportControl.md)|
|[DFirst](../../../api/Access.Application.DFirst.md)|
|[DirtyObject](../../../api/Access.Application.DirtyObject.md)|
|[DLast](../../../api/Access.Application.DLast.md)|
|[DLookup](../../../api/Access.Application.DLookup.md)|
|[DMax](../../../api/Access.Application.DMax.md)|
|[DMin](../../../api/Access.Application.DMin.md)|
|[DStDev](../../../api/Access.Application.DStDev.md)|
|[DStDevP](../../../api/Access.Application.DStDevP.md)|
|[DSum](../../../api/Access.Application.DSum.md)|
|[DVar](../../../api/Access.Application.DVar.md)|
|[DVarP](../../../api/Access.Application.DVarP.md)|
|[Echo](../../../api/Access.Application.Echo.md)|
|[EuroConvert](../../../api/Access.Application.EuroConvert.md)|
|[Eval](../../../api/Access.Application.Eval.md)|
|[ExportNavigationPane](../../../api/Access.Application.ExportNavigationPane.md)|
|[ExportXML](../../../api/Access.Application.ExportXML.md)|
|[FollowHyperlink](../../../api/Access.Application.FollowHyperlink.md)|
|[GetHiddenAttribute](../../../api/Access.Application.GetHiddenAttribute.md)|
|[GetOption](../../../api/Access.Application.GetOption.md)|
|[GUIDFromString](../../../api/Access.Application.GUIDFromString.md)|
|[HtmlEncode](../../../api/Access.Application.HtmlEncode.md)|
|[hWndAccessApp](../../../api/Access.Application.hWndAccessApp.md)|
|[HyperlinkPart](../../../api/Access.Application.HyperlinkPart.md)|
|[ImportNavigationPane](../../../api/Access.Application.ImportNavigationPane.md)|
|[ImportXML](../../../api/Access.Application.ImportXML.md)|
|[InstantiateTemplate](../../../api/Access.Application.InstantiateTemplate.md)|
|[IsCurrentWebUserInGroup](../../../api/Access.Application.IsCurrentWebUserInGroup.md)|
|[LoadCustomUI](../../../api/Access.Application.LoadCustomUI.md)|
|[LoadFromAXL](../../../api/Access.Application.LoadFromAXL.md)|
|[LoadPicture](../../../api/Access.Application.LoadPicture.md)|
|[NewAccessProject](../../../api/Access.Application.NewAccessProject.md)|
|[NewCurrentDatabase](../../../api/Access.Application.NewCurrentDatabase.md)|
|[Nz](../../../api/Access.Application.Nz.md)|
|[OpenAccessProject](../../../api/Access.Application.OpenAccessProject.md)|
|[OpenCurrentDatabase](../../../api/Access.Application.OpenCurrentDatabase.md)|
|[PlainText](../../../api/Access.Application.PlainText.md)|
|[Quit](../../../api/Access.Application.Quit.md)|
|[RefreshDatabaseWindow](../../../api/Access.Application.RefreshDatabaseWindow.md)|
|[RefreshTitleBar](../../../api/Access.Application.RefreshTitleBar.md)|
|[Run](../../../api/Access.Application.Run.md)|
|[RunCommand](../../../api/Access.Application.RunCommand.md)|
|[SaveAsAXL](../../../api/Access.Application.SaveAsAXL.md)|
|[SaveAsTemplate](../../../api/Access.Application.SaveAsTemplate.md)|
|[SetDefaultWorkgroupFile](../../../api/Access.Application.SetDefaultWorkgroupFile.md)|
|[SetHiddenAttribute](../../../api/Access.Application.SetHiddenAttribute.md)|
|[SetOption](../../../api/Access.Application.SetOption.md)|
|[StringFromGUID](../../../api/Access.Application.StringFromGUID.md)|
|[SysCmd](../../../api/Access.Application.SysCmd.md)|
|[TransformXML](../../../api/Access.Application.TransformXML.md)|

## Properties



|**Name**|
|:-----|
|[Application](../../../api/Access.Application.Application.md)|
|[Assistance](../../../api/Access.Application.Assistance.md)|
|[AutoCorrect](../../../api/Access.Application.AutoCorrect.md)|
|[AutomationSecurity](../../../api/Access.Application.AutomationSecurity.md)|
|[BrokenReference](../../../api/Access.Application.BrokenReference.md)|
|[Build](../../../api/Access.Application.Build.md)|
|[CodeContextObject](../../../api/Access.Application.CodeContextObject.md)|
|[CodeData](../../../api/Access.Application.CodeData.md)|
|[CodeProject](../../../api/Access.Application.CodeProject.md)|
|[COMAddIns](../../../api/Access.Application.COMAddIns.md)|
|[CommandBars](../../../api/Access.Application.CommandBars.md)|
|[CurrentData](../../../api/Access.Application.CurrentData.md)|
|[CurrentObjectName](../../../api/Access.Application.CurrentObjectName.md)|
|[CurrentObjectType](../../../api/Access.Application.CurrentObjectType.md)|
|[CurrentProject](../../../api/Access.Application.CurrentProject.md)|
|[DBEngine](../../../api/Access.Application.DBEngine.md)|
|[DoCmd](../../../api/Access.Application.DoCmd.md)|
|[FeatureInstall](../../../api/Access.Application.FeatureInstall.md)|
|[FileDialog](../../../api/Access.Application.FileDialog.md)|
|[Forms](../../../api/Access.Application.Forms.md)|
|[IsCompiled](../../../api/Access.Application.IsCompiled.md)|
|[LanguageSettings](../../../api/Access.Application.LanguageSettings.md)|
|[MacroError](../../../api/Access.Application.MacroError.md)|
|[MenuBar](../../../api/Access.Application.MenuBar.md)|
|[Modules](../../../api/Access.Application.Modules.md)|
|[Name](../../../api/Access.Application.Name.md)|
|[NewFileTaskPane](../../../api/Access.Application.NewFileTaskPane.md)|
|[Parent](../../../api/Access.Application.Parent.md)|
|[Printer](../../../api/Access.Application.Printer.md)|
|[Printers](../../../api/Access.Application.Printers.md)|
|[ProductCode](../../../api/Access.Application.ProductCode.md)|
|[References](../../../api/Access.Application.References.md)|
|[Reports](../../../api/Access.Application.Reports.md)|
|[ReturnVars](../../../api/Access.application.returnvars.md)|
|[Screen](../../../api/Access.Application.Screen.md)|
|[ShortcutMenuBar](../../../api/Access.Application.ShortcutMenuBar.md)|
|[TempVars](../../../api/Access.Application.TempVars.md)|
|[UserControl](../../../api/Access.Application.UserControl.md)|
|[VBE](../../../api/Access.Application.VBE.md)|
|[Version](../../../api/Access.Application.Version.md)|
|[Visible](../../../api/Access.Application.Visible.md)|
|[WebServices](../../../api/Access.Application.WebServices.md)|

## See also

[Access Object Model Reference](object-model-access-vba-reference.md)


