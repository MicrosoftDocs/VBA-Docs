---
title: CustomTaskPane object (Office)
keywords: vbaof11.chm3030000
f1_keywords:
- vbaof11.chm3030000
ms.prod: office
api_name:
- Office.CustomTaskPane
ms.assetid: 7ed379b7-d070-4d7b-abe1-92dc73d3d137
ms.date: 01/04/2019
localization_priority: Normal
---


# CustomTaskPane object (Office)

Represents a custom task pane in the container application.


## Example

The following example, written in C#, creates an instance of a **CustomTaskPane** object and implements its only method, **CTPFactoryAvailable**. **CTPFactoryAvailable** passes an **ICTPFactory** object to the add-in, which you can use during the add-in's lifetime to create a task pane by using the **CreateCTP** method. Note that the example assumes that the task pane is part of a COM add-in and thus implements **Extensibility.IDTExtensibility2**. The add-in also references a Microsoft ActiveX control, **SampleActiveX.myControl**, that was created in a separate project.


```cs
public class Connect : Object, Extensibility.IDTExtensibility2, ICustomTaskPaneConsumer 
... 
object missing = Type.Missing; 
public CustomTaskPane CTP = null; 
 
public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst) 
{ 
 CTP = CTPFactoryInst.CreateCTP("SampleActiveX.myControl", "Task Pane Example", missing); 
 sampleAX = (myControl)CTP.ContentControl; 
 sampleAX.InsertTextClicked += new InsertTextEventHandler(sampleAX_InsertTextClicked); 
 CTP.Visible = true; 
} 
...
```


> [!NOTE] 
> You can create custom task panes in any language that supports COM and allows you to create dynamic-linked library (DLL) files; for example, Microsoft Visual Basic 6.0, Visual Basic .NET, Visual C++, Visual C++ .NET, and Visual C#. However, Microsoft Visual Basic for Applications (VBA) does not support creating custom task panes. 


## See also

- [CustomTaskPane object members](overview/library-reference/customtaskpane-members-office.md)
- [Object Model Reference](overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]