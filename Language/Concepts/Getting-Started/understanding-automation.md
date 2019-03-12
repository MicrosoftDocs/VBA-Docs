---
title: Understanding automation (VBA)
keywords: vbcn6.chm1076677
f1_keywords:
- vbcn6.chm1076677
ms.prod: office
ms.assetid: 5b45f6f3-1459-ff25-51e1-32c475f11153
ms.date: 12/21/2018
localization_priority: Normal
---


# Understanding automation

Automation (formerly OLE Automation) is a feature of the Component Object Model (COM), an industry-standard technology that applications use to expose their [objects](../../Glossary/vbe-glossary.md#object) to development tools, macro languages, and other applications that support Automation. For example, a spreadsheet application may expose a worksheet, chart, cell, or range of cells—each as a different type of object. A word processor might expose objects such as an application, a document, a paragraph, a sentence, a bookmark, or a selection.

When an application supports Automation, the objects the application exposes can be accessed by Visual Basic. Use Visual Basic to manipulate these objects by invoking [methods](../../Glossary/vbe-glossary.md#method) on the object or by getting and setting the object's properties. For example, you can create an [Automation object](../../Glossary/vbe-glossary.md#automation-object) and write the following code to access the object.

```vb
MyObj.Insert "Hello, world." ' Place text. 
MyObj.Bold = True ' Format text. 
If Mac = True ' Check your platform constant 
 MyObj.SaveAs "HD:\WORDPROC\DOCS\TESTOBJ.DOC" ' Save the object (Macintosh). 
Else 
 MyObj.SaveAs "C:\WORDPROC\DOCS\TESTOBJ.DOC" ' Save the object (Windows). 

```

<br/>

Use the following functions to access an Automation object.

|Function|Description|
|:-----|:-----|
|**[CreateObject](../../reference/user-interface-help/createobject-function.md)**|Creates a new object of a specified type.|
|**[GetObject](../../reference/user-interface-help/getobject-function.md)**|Retrieves an object from a file.|

For information about the properties and methods supported by an application, see the application documentation. The objects, functions, properties, and methods supported by an application are usually defined in the application's [object library](../../Glossary/vbe-glossary.md#object-library).

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
