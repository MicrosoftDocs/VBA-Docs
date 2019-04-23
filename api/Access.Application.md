---
title: Application object (Access)
keywords: vbaac10.chm12627
f1_keywords:
- vbaac10.chm12627
ms.prod: access
api_name:
- Access.Application
ms.assetid: aefb0713-97e6-e2c7-e530-8fd2e1316a55
ms.date: 02/05/2019
localization_priority: Normal
---


# Application object (Access)

The **Application** object refers to the active Microsoft Access application.

## Remarks

The **Application** object contains all Access objects and collections.

You can use the **Application** object to apply methods or property settings to the entire Access application. For example, you can use the **SetOption** method of the **Application** object to set database options from Visual Basic. The following example shows how you can set the **Display Status Bar** check box on the **Current Database** tab of the **Access Options** dialog box.

```vb
Application.SetOption "Show Status Bar", True
```

Access is a COM component that supports Automation, formerly called OLE Automation. You can manipulate Access objects from another application that also supports Automation. To do this, you use the **Application** object.

For example, Microsoft Visual Basic is a COM component. You can open an Access database from Visual Basic and work with its objects. From Visual Basic, first create a reference to the Access object library, and then create a new instance of the **Application** class and point an object variable to it, as in the following example:

```vb
Dim appAccess As New Access.Application
```

From applications that do not support the **New** keyword, you can create a new instance of the **Application** class by using the **CreateObject** function:

```vb
Dim appAccess As Object 
Set appAccess = CreateObject("Access.Application")
```

After you create a new instance of the **Application** class, you can open a database or create a new database by using either the **OpenCurrentDatabase** method or the **NewCurrentDatabase** method. You can then set the properties of the **Application** object and call its methods. 

When you return a reference to the **CommandBars** object by using the **CommandBars** property of the **Application** object, you can access all Microsoft Office command bar objects and collections by using this reference.

You can also manipulate other Access objects through the **Application** object. For example, by using the **[OpenForm](Access.DoCmd.OpenForm.md)** method of the Access **[DoCmd](Access.DoCmd.md)** object, you can open an Access form from Microsoft Office Excel:

```vb
appAccess.DoCmd.OpenForm "Orders"
```

For more information about creating a reference and controlling objects by using Automation, see the documentation for the application that is acting as the COM component.


## Methods

- [AccessError](Access.Application.AccessError.md)
- [AddToFavorites](Access.Application.AddToFavorites.md)
- [BuildCriteria](Access.Application.BuildCriteria.md)
- [CloseCurrentDatabase](Access.Application.CloseCurrentDatabase.md)
- [CodeDb](Access.Application.CodeDb.md)
- [ColumnHistory](Access.Application.ColumnHistory.md)
- [ConvertAccessProject](Access.Application.ConvertAccessProject.md)
- [CreateAccessProject](Access.Application.CreateAccessProject.md)
- [CreateAdditionalData](Access.Application.CreateAdditionalData.md)
- [CreateControl](Access.Application.CreateControl.md)
- [CreateForm](Access.Application.CreateForm.md)
- [CreateGroupLevel](Access.Application.CreateGroupLevel.md)
- [CreateReport](Access.Application.CreateReport.md)
- [CreateReportControl](Access.Application.CreateReportControl.md)
- [CurrentDb](Access.Application.CurrentDb.md)
- [CurrentUser](Access.Application.CurrentUser.md)
- [CurrentWebUser](Access.Application.CurrentWebUser.md)
- [CurrentWebUserGroups](Access.Application.CurrentWebUserGroups.md)
- [DAvg](Access.application.davg.md)
- [DCount](Access.Application.DCount.md)
- [DDEExecute](Access.Application.DDEExecute.md)
- [DDEInitiate](Access.Application.DDEInitiate.md)
- [DDEPoke](Access.Application.DDEPoke.md)
- [DDERequest](Access.Application.DDERequest.md)
- [DDETerminate](Access.Application.DDETerminate.md)
- [DDETerminateAll](Access.Application.DDETerminateAll.md)
- [DefaultWorkspaceClone](Access.Application.DefaultWorkspaceClone.md)
- [DeleteControl](Access.Application.DeleteControl.md)
- [DeleteReportControl](Access.Application.DeleteReportControl.md)
- [DFirst](Access.Application.DFirst.md)
- [DirtyObject](Access.Application.DirtyObject.md)
- [DLast](Access.Application.DLast.md)
- [DLookup](Access.Application.DLookup.md)
- [DMax](Access.Application.DMax.md)
- [DMin](Access.Application.DMin.md)
- [DStDev](Access.Application.DStDev.md)
- [DStDevP](Access.Application.DStDevP.md)
- [DSum](Access.Application.DSum.md)
- [DVar](Access.Application.DVar.md)
- [DVarP](Access.Application.DVarP.md)
- [Echo](Access.Application.Echo.md)
- [EuroConvert](Access.Application.EuroConvert.md)
- [Eval](Access.Application.Eval.md)
- [ExportNavigationPane](Access.Application.ExportNavigationPane.md)
- [ExportXML](Access.Application.ExportXML.md)
- [FollowHyperlink](Access.Application.FollowHyperlink.md)
- [GetHiddenAttribute](Access.Application.GetHiddenAttribute.md)
- [GetOption](Access.Application.GetOption.md)
- [GUIDFromString](Access.Application.GUIDFromString.md)
- [HtmlEncode](Access.Application.HtmlEncode.md)
- [hWndAccessApp](Access.Application.hWndAccessApp.md)
- [HyperlinkPart](Access.Application.HyperlinkPart.md)
- [ImportNavigationPane](Access.Application.ImportNavigationPane.md)
- [ImportXML](Access.Application.ImportXML.md)
- [InstantiateTemplate](Access.Application.InstantiateTemplate.md)
- [IsCurrentWebUserInGroup](Access.Application.IsCurrentWebUserInGroup.md)
- [LoadCustomUI](Access.Application.LoadCustomUI.md)
- [LoadFromAXL](Access.Application.LoadFromAXL.md)
- [LoadPicture](Access.Application.LoadPicture.md)
- [NewAccessProject](Access.Application.NewAccessProject.md)
- [NewCurrentDatabase](Access.Application.NewCurrentDatabase.md)
- [Nz](Access.Application.Nz.md)
- [OpenAccessProject](Access.Application.OpenAccessProject.md)
- [OpenCurrentDatabase](Access.Application.OpenCurrentDatabase.md)
- [PlainText](Access.Application.PlainText.md)
- [Quit](Access.Application.Quit.md)
- [RefreshDatabaseWindow](Access.Application.RefreshDatabaseWindow.md)
- [RefreshTitleBar](Access.Application.RefreshTitleBar.md)
- [Run](Access.Application.Run.md)
- [RunCommand](Access.Application.RunCommand.md)
- [SaveAsAXL](Access.Application.SaveAsAXL.md)
- [SaveAsTemplate](Access.Application.SaveAsTemplate.md)
- [SetDefaultWorkgroupFile](Access.Application.SetDefaultWorkgroupFile.md)
- [SetHiddenAttribute](Access.Application.SetHiddenAttribute.md)
- [SetOption](Access.Application.SetOption.md)
- [StringFromGUID](Access.Application.StringFromGUID.md)
- [SysCmd](Access.Application.SysCmd.md)
- [TransformXML](Access.Application.TransformXML.md)



## Properties

- [AppIcon](Access.Application.AppIcon.md)
- [Application](Access.Application.Application.md)
- [AppTitle](Access.Application.AppTitle.md)
- [Assistance](Access.Application.Assistance.md)
- [AutoCorrect](Access.Application.AutoCorrect.md)
- [AutomationSecurity](Access.Application.AutomationSecurity.md)
- [BrokenReference](Access.Application.BrokenReference.md)
- [Build](Access.Application.Build.md)
- [CodeContextObject](Access.Application.CodeContextObject.md)
- [CodeData](Access.Application.CodeData.md)
- [CodeProject](Access.Application.CodeProject.md)
- [COMAddIns](Access.Application.COMAddIns.md)
- [CommandBars](Access.Application.CommandBars.md)
- [CurrentData](Access.Application.CurrentData.md)
- [CurrentObjectName](Access.Application.CurrentObjectName.md)
- [CurrentObjectType](Access.Application.CurrentObjectType.md)
- [CurrentProject](Access.Application.CurrentProject.md)
- [DBEngine](Access.Application.DBEngine.md)
- [DoCmd](Access.Application.DoCmd.md)
- [FeatureInstall](Access.Application.FeatureInstall.md)
- [FileDialog](Access.Application.FileDialog.md)
- [Forms](Access.Application.Forms.md)
- [IsCompiled](Access.Application.IsCompiled.md)
- [LanguageSettings](Access.Application.LanguageSettings.md)
- [MacroError](Access.Application.MacroError.md)
- [MenuBar](Access.Application.MenuBar.md)
- [Modules](Access.Application.Modules.md)
- [Name](Access.Application.Name.md)
- [NewFileTaskPane](Access.Application.NewFileTaskPane.md)
- [Parent](Access.Application.Parent.md)
- [Printer](Access.Application.Printer.md)
- [Printers](Access.Application.Printers.md)
- [ProductCode](Access.Application.ProductCode.md)
- [References](Access.Application.References.md)
- [Reports](Access.Application.Reports.md)
- [ReturnVars](Access.application.returnvars.md)
- [Screen](Access.Application.Screen.md)
- [ShortcutMenuBar](Access.Application.ShortcutMenuBar.md)
- [TempVars](Access.Application.TempVars.md)
- [UserControl](Access.Application.UserControl.md)
- [VBE](Access.Application.VBE.md)
- [Version](Access.Application.Version.md)
- [Visible](Access.Application.Visible.md)
- [WebServices](Access.Application.WebServices.md)

## See also

- [Access Object Model reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
