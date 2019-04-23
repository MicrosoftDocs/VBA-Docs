---
title: AccessObject object (Access)
keywords: vbaac10.chm12743
f1_keywords:
- vbaac10.chm12743
ms.prod: access
api_name:
- Access.AccessObject
ms.assetid: 8a770b33-5bff-120a-6707-ca214ee5ced3
ms.date: 02/01/2019
localization_priority: Normal
---


# AccessObject object (Access)

An **AccessObject** object refers to a particular Access object.


## Remarks

An **AccessObject** object includes information about one instance of an object. The following table list the types of objects each **AccessObject** describes, the name of its collection, and what type of information **AccessObject** contains.

<br/>

|AccessObject|Collection|Contains information about|
|:-----|:-----|:-----|
|**Database diagram**|**AllDatabaseDiagrams**|Saved database diagrams|
|**Form**|**AllForms**|Saved forms|
|**Function**|**AllFunctions**|Saved functions|
|**Macro**|**AllMacros**|Saved macros|
|**Module**|**AllModules**|Saved modules|
|**Query**|**AllQueries**|Saved queries|
|**Report**|**AllReports**|Saved reports|
|**Stored procedure**|**AllStoredProcedures**|Saved stored procedures|
|**Table**|**AllTables**|Saved tables|
|**View**|**AllViews**|Saved views|

Because an **AccessObject** object corresponds to an existing object, you can't create new **AccessObject** objects or delete existing ones. To refer to an **AccessObject** object in a collection by its ordinal number or by its **Name** property setting, use any of the following syntax forms:

- **AllForms** (0)
- **AllForms** ("_name_")
- **AllForms** ![ _name_ ]

## Methods

- [GetDependencyInfo](Access.AccessObject.GetDependencyInfo.md)
- [IsDependentUpon](Access.AccessObject.IsDependentUpon.md)

## Properties

- [CurrentView](Access.AccessObject.CurrentView.md)
- [DateCreated](Access.AccessObject.DateCreated.md)
- [DateModified](Access.AccessObject.DateModified.md)
- [FullName](Access.AccessObject.FullName.md)
- [IsLoaded](Access.AccessObject.IsLoaded.md)
- [IsWeb](Access.AccessObject.IsWeb.md)
- [Name](Access.AccessObject.Name.md)
- [Parent](Access.AccessObject.Parent.md)
- [Properties](Access.AccessObject.Properties.md)
- [Type](Access.AccessObject.Type.md)

## See also

- [Access Object Model Reference](overview/access/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
