---
title: ICustomTaskPaneConsumer object (Office)
keywords: vbaof11.chm305000
f1_keywords:
- vbaof11.chm305000
ms.prod: office
api_name:
- Office.ICustomTaskPaneConsumer
ms.assetid: 54be3f78-4e5d-8595-d369-0724df0debf7
ms.date: 01/16/2019
localization_priority: Normal
---


# ICustomTaskPaneConsumer object (Office)

An interface that provides access to the **[CTPFactoryAvailable](Office.ICustomTaskPaneConsumer.CTPFactoryAvailable.md)** method that is used to create an instance of a custom task pane.


## Example

The following example, written in C#, creates an instance of a **CustomTaskPane** object through the **ICustomTaskPaneConsumer** interface and implements its only method, **CTPFactoryAvailable**. **CTPFactoryAvailable** passes an **ICTPFactory** object to the add-in, which you can use during the add-in's lifetime to create task panes by using the **CreateCTP** method. 

Note that the example assumes that the task pane is part of a COM add-in and thus implements **Extensibility.IDTExtensibility2**. The add-in also references a Microsoft ActiveX control, SampleActiveX.myControl, which is created in a separate project.


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
> You can create custom task panes in any language that supports COM and allows you to create dynamic-linked library (DLL) files; for example, Microsoft Visual Basic 6.0, Visual Basic .NET, Visual C++, Visual C++ .NET, and Visual C#. However, Visual Basic for Applications (VBA) does not support creating custom task panes. 


## See also

- [ICustomTaskPaneConsumer object members](overview/Library-Reference/icustomtaskpaneconsumer-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]