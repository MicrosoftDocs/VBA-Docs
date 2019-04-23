---
title: Calling procedures with the same name (VBA)
keywords: vbcn6.chm1076672
f1_keywords:
- vbcn6.chm1076672
ms.prod: office
ms.assetid: 5d310675-136b-58bb-29e2-ca09726b8ce0
ms.date: 12/21/2018
localization_priority: Normal
---


# Calling procedures with the same name

You can call a [procedure](../../Glossary/vbe-glossary.md#procedure) located in any [module](../../Glossary/vbe-glossary.md#module) in the same [project](../../Glossary/vbe-glossary.md#project) as the active module just as you would call a procedure in the active module. However, if two or more modules contain a procedure with the same name, you must specify a module name in the calling statement, as shown in the following example:

```vb
Sub Main() 
    Module1.MyProcedure 
End Sub
```

<br/>

If you give the same name to two different procedures in two different projects, you must specify a project name when you call that procedure. For example, the following procedure calls the `Main` procedure in the `MyModule` module in the `MyProject.vbp` project.

```vb
Sub Main() 
    [MyProject.vbp].[MyModule].Main 
End Sub
```

> [!NOTE] 
> Different applications have different names for a project. For example, in Microsoft Access, a project is called a database (.mdb); in Microsoft Excel, it's a workbook (.xls).


> [!TIP] 
> - If you rename a module or project, be sure to change the module or project name wherever it appears in calling [statements](../../Glossary/vbe-glossary.md#statement); otherwise, Visual Basic will not be able to find the called procedure. You can use the **Replace** command on the **[Edit](../../reference/user-interface-help/edit-menu.md)** menu to find and replace text in a module.
> - To avoid naming conflicts among referenced projects, give your procedures unique names so you can call a procedure without specifying a project or module.
    
## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
