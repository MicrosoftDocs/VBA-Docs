---
title: CustomTaskPane.Delete method (Office)
keywords: vbaof11.chm301010
f1_keywords:
- vbaof11.chm301010
ms.prod: office
api_name:
- Office.CustomTaskPane.Delete
ms.assetid: 6db4b7ba-3dd8-7249-07dc-511516b1a16c
ms.date: 01/04/2019
localization_priority: Normal
---


# CustomTaskPane.Delete method (Office)

Deletes the active custom task pane.


## Syntax

_expression_.**Delete**

_expression_ An expression that returns a **[CustomTaskPane](Office.CustomTaskPane.md)** object.


## Example

The following example, written in C#, creates an instance of a **CustomTaskPane** object and implements its only method, **CTPFactoryAvailable**. **CTPFactoryAvailable** passes a **CTPFactory** object to the add-in, which can be used during the add-in's lifetime to create task panes by using the **CreateCTP** method. The project also implements a button that is used to delete the active task pane. Note that the example assumes that the task pane is part of a COM add-in and thus implements **Extensibility.IDTExtensibility2**. The add-in also refers to a Microsoft ActiveX control, **SampleActiveX.myControl**, that was created in a separate project.


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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]