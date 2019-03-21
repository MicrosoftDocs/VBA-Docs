---
title: Module object (Access)
keywords: vbaac10.chm12268
f1_keywords:
- vbaac10.chm12268
ms.prod: access
api_name:
- Access.Module
ms.assetid: e04272fa-9c29-2567-bd15-1cea38906894
ms.date: 03/21/2019
localization_priority: Normal
---


# Module object (Access)

A **Module** object refers to a standard module or a class module.


## Remarks

Microsoft Access includes class modules that are not associated with any object, and form modules and report modules, which are associated with a form or report.

To determine whether a **Module** object represents a standard module or a class module from code, check the **Module** object's **Type** property.

The **[Modules](Access.Modules.md)** collection contains all open **Module** objects, regardless of their type. Modules in the **Modules** collection can be compiled or uncompiled.

To return a reference to a particular standard or class **Module** object in the **Modules** collection, use any of the following syntax forms.

|Syntax|Description|
|:-----|:-----|
|**Modules**!_modulename_|The _modulename_ argument is the name of the **Module** object.|
|**Modules**("_modulename_")|The _modulename_ argument is the name of the **Module** object.|
|**Modules**(_index_)|The _index_ argument is the numeric position of the object within the collection.|

The following example returns a reference to a standard **Module** object and assigns it to an object variable.

```vb
Dim mdl As Module 
Set mdl = Modules![Utility Functions]
```

Note that the brackets enclosing the name of the **Module** object are necessary only if the name of the **Module** object includes spaces.

The next example returns a reference to a form **Module** object and assigns it to an object variable.

```vb
Dim mdl As Module 
Set mdl = Modules!Form_Employees
```

To refer to a specific form or report module, you can also use the **[Form](Access.Form.md)** or **[Report](Access.Report.md)** object's **Module** property.

```vb
Forms!formname .Module
```

The following example also returns a reference to the **Module** object associated with an **Employees** form and assigns it to an object variable.

```vb
Dim mdl As Module 
Set mdl = Forms!Employees.Module
```

After you've returned a reference to a **Module** object, you can set or read its properties and apply its methods.


## Methods

- [AddFromFile](Access.Module.AddFromFile.md)
- [AddFromString](Access.Module.AddFromString.md)
- [CreateEventProc](Access.Module.CreateEventProc.md)
- [DeleteLines](Access.Module.DeleteLines.md)
- [Find](Access.Module.Find.md)
- [InsertLines](Access.Module.InsertLines.md)
- [InsertText](Access.Module.InsertText.md)
- [ReplaceLine](Access.Module.ReplaceLine.md)

## Properties

- [Application](Access.Module.Application.md)
- [CountOfDeclarationLines](Access.Module.CountOfDeclarationLines.md)
- [CountOfLines](Access.Module.CountOfLines.md)
- [Lines](Access.Module.Lines.md)
- [Name](Access.Module.Name.md)
- [Parent](Access.Module.Parent.md)
- [ProcBodyLine](Access.Module.ProcBodyLine.md)
- [ProcCountLines](Access.Module.ProcCountLines.md)
- [ProcOfLine](Access.Module.ProcOfLine.md)
- [ProcStartLine](Access.Module.ProcStartLine.md)
- [Type](Access.Module.Type.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
