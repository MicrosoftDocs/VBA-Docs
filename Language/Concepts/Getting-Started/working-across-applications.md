---
title: Working across applications (VBA)
keywords: vbcn6.chm1012581
f1_keywords:
- vbcn6.chm1012581
ms.prod: office
ms.assetid: 46d31003-fdfb-04d3-6143-e286d91a46a8
ms.date: 12/26/2018
localization_priority: Normal
---


# Working across applications

Visual Basic can create new [objects](../../Glossary/vbe-glossary.md#object) and retrieve existing objects from many Microsoft applications. Other applications may also provide objects that you can create by using Visual Basic. See the application's documentation for more information.

To create a new object or get an existing object from another application, use the **[CreateObject](../../reference/user-interface-help/createobject-function.md)** function or **[GetObject](../../reference/user-interface-help/getobject-function.md)** function.

```vb
' Start Microsoft Excel and create a new Worksheet object. 
Set ExcelWorksheet = CreateObject("Excel.Sheet") 
 
' Start Microsoft Excel and open an existing Worksheet object. 
Set ExcelWorksheet = GetObject("SHEET1.XLS") 
 
' Start Microsoft Word. 
Set WordBasic = CreateObject("Word.Basic") 

```

Most applications provide an **Exit** or **Quit** method that closes the application whether or not it is visible. For more information about the objects, methods, and properties an application provides, see the application's documentation.

Some applications allow you to use the **New** [keyword](../../Glossary/vbe-glossary.md#keyword) to create an object of any class that exists in its [type library](../../Glossary/vbe-glossary.md#type-library). For example:

```vb
Dim X As New Field 

```

This case is an example of a [class](../../Glossary/vbe-glossary.md#class) in the data access type library. A new instance of a **Field** object is created by using this syntax. Refer to the application's documentation for information about which object classes can be created in this way.

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]