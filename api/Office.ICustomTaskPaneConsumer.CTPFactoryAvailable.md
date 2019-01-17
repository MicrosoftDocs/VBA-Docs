---
title: ICustomTaskPaneConsumer.CTPFactoryAvailable method (Office)
keywords: vbaof11.chm305001
f1_keywords:
- vbaof11.chm305001
ms.prod: office
api_name:
- Office.ICustomTaskPaneConsumer.CTPFactoryAvailable
ms.assetid: b4fd5ea5-5cad-0c48-0538-855f94fb65c9
ms.date: 01/16/2019
localization_priority: Normal
---


# ICustomTaskPaneConsumer.CTPFactoryAvailable method (Office)

Passes an **[ICTPFactory](office.ictpfactory.md)** object to a Microsoft ActiveX add-in that can then be used when creating a custom task pane.


## Syntax

_expression_.**CTPFactoryAvailable** (_CTPFactoryInst_)

_expression_ An expression that returns an **[ICustomTaskPaneConsumer](Office.ICustomTaskPaneConsumer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _CTPFactoryInst_|Required|**ICTPFactory**|The object is used by an add-in to create a task pane.|

## Example

The following example, written in C#, creates an instance of a **CustomTaskPane** object through the **ICustomTaskPaneConsumer** interface and implements its only method, **CTPFactoryAvailable**. **CTPFactoryAvailable** passes an **ICTPFactory** object to the add-in, which you can use during the add-in's lifetime to create a task pane by using the **CreateCTP** method. 

Note that the example assumes that the task pane is part of a COM add-in and thus implements **Extensibility.IDTExtensibility2**. The add-in also refers to an ActiveX control, SampleActiveX.myControl, which was created in a separate project.


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
```

> [!NOTE] 
> You can create custom task panes in any language that supports COM and allows you to create dynamic-linked library (DLL) files; for example, Microsoft Visual Basic 6.0, Visual Basic .NET, Visual C++, Visual C++ .NET, and Visual C#. However, Visual Basic for Applications (VBA) does not support creating custom task panes. 



## See also

- [ICustomTaskPaneConsumer object members](overview/Library-Reference/icustomtaskpaneconsumer-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]