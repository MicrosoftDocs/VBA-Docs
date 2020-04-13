---
title: AccessObjectProperty object (Access)
keywords: vbaac10.chm12693
f1_keywords:
- vbaac10.chm12693
ms.prod: access
api_name:
- Access.AccessObjectProperty
ms.assetid: b1a44d34-8ca1-af7d-1878-f2c14fb481f7
ms.date: 02/01/2019
localization_priority: Normal
---


# AccessObjectProperty object (Access)

An **AccessObjectProperty** object represents a built-in or user-defined characteristic of an **[AccessObject](access.accessobject.md)** object.


## Remarks

Every **AccessObject** object contains an **[AccessObjectProperties](access.accessobjectproperties.md)** collection that has **AccessObjectProperty** objects corresponding to the properties of that **AccessObject** object. The user can also define **AccessObjectProperty** objects and append them to the **AccessObjectProperties** collection of some **AccessObject** objects.

You can create user-defined properties for the following objects:

- **CodeData**, **CodeProject**, **CurrentProject**, and **CurrentData** objects

- **AccessObject** objects in the following collections:

  - CurrentProject and CodeProject object collections:

    - **[AllForms](Access.AllForms.md)**
    - **[AllReports](Access.AllReports.md)**
    - **[AllMacros](Access.allmacros.md)**
    - **[AllModules](Access.AllModules.md)**
    - **[AllTables](Access.AllTables.md)**

  - CodeData and CodeProject object collections:

    - **[AllQueries](Access.AllQueries.md)**
    - **[AllViews](Access.AllViews.md)**
    - **[AllStoredProcedures](Access.AllStoredProcedures.md)**
    - **[AllDatabaseDiagrams](Access.AllDatabaseDiagrams.md)**

> [!NOTE]
> The **AccessObjectProperties** collection isn't accessible for objects derived from the **[CurrentData](access.currentdata.md)** object (for example, CurrentData.AllTables!Table1). For objects derived in this manner, you can only access their built-in properties by direct calls to the desired property (for example, CurrentData.AllTables!Table1.Name).

To add a user-defined property, use the **[Add](Access.AccessObjectProperties.Add.md)** method to create and add an **AccessObjectProperty** object with a unique **Name** property and **Value** property. The object to which you are adding the user-defined property must already be appended to a collection. 

Referencing a user-defined **AccessObjectProperty** object that has not yet been appended to an **AccessObjectProperties** collection will cause an error, as will appending a user-defined **AccessObjectProperty** object to an **AccessObjectProperties** collection containing an **AccessObjectProperty** object of the same name.

You can delete user-defined properties from the **AccessObjectProperties** collection by using the **[Remove](Access.AccessObjectProperties.Remove.md)** method.

> [!NOTE] 
> A user-defined **AccessObjectProperty** object is associated only with the specific instance of an object. The property isn't defined for all instances of objects of the selected type.

The **AccessObjectProperty** object has two built-in properties:

- The **Name** property, a **String** that uniquely identifies the property.
- The **Value** property, a **Variant** that contains the property setting.

To refer to a built-in or user-defined **AccessObjectProperty** object in a collection by its ordinal number or by its **Name** property setting, use any of the following syntax forms.

```vb
CurrentProject.AllForms("Form1").Properties(0) 
CurrentProject.AllForms("Form1").Properties("name") 
CurrentProject.AllForms("Form1").Properties![name]
```

With the same syntax forms, you can also refer to the **Value** property of an **AccessObjectProperty** object. The context of the reference will determine whether you are referring to the **AccessObjectProperty** object itself or to the **Value** property of the **AccessObjectProperty** object.

> [!NOTE] 
> Properties in the **AccessObjectProperties** collection are not stored and can be lost when the object they are associated with is checked in or out by using the **Source Code Control** add-in.


## Properties

- [Name](Access.AccessObjectProperty.Name.md)
- [Value](Access.AccessObjectProperty.Value.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]