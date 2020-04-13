---
title: Using events with the Application object (Word)
keywords: vbawd10.chm5214014
f1_keywords:
- vbawd10.chm5214014
ms.prod: word
ms.assetid: 784c4c61-7e47-3dbf-46f6-da655f786ca1
ms.date: 06/08/2019
localization_priority: Normal
---


# Using events with the Application object (Word)

To create an event handler for an event of the **[Application](../../../api/Word.Application.md)** object, you need to complete the following three steps:


1. Declare an object variable in a class module to respond to the events.
    
2. Write the specific event procedures.
    
3. Initialize the declared object from another module.
    

## Declare the object variable

Before you can write procedures for the events of the **Application** object, you must create a new class module and declare an object of type **Application** with events. For example, assume that a new class module is created and called EventClassModule. The new class module contains the following code.


```vb
Public WithEvents App As Word.Application
```


## Write the event procedures

After the new object has been declared with events, it appears in the **Object** drop-down list box in the class module, and you can write event procedures for the new object. (When you select the new object in the **Object** box, the valid events for that object are listed in the **Procedure** drop-down list box.) Select an event from the **Procedure** drop-down list box; an empty procedure is added to the class module.


```vb
Private Sub App_DocumentChange() 
 
End Sub
```


## Initialize the declared object

Before the procedure will run, you must connect the declared object in the class module (App in this example) with the **Application** object. You can do this with the following code from any module.


```vb
Dim X As New EventClassModule 
Sub Register_Event_Handler() 
 Set X.App = Word.Application 
End Sub
```

Run the Register_Event_Handler procedure. After the procedure is run, the App object in the class module points to the Microsoft Word **Application** object, and the event procedures in the class module will run when the events occur.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
