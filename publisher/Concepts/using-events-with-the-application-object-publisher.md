---
title: Using events with the Application object (Publisher)
ms.prod: publisher
ms.assetid: 29b60d3c-3049-2ba9-8688-e46c4323e9ba
ms.date: 06/04/2019
localization_priority: Normal
---


# Using events with the Application object (Publisher)

To create an event handler for an event of the **[Application](../../api/publisher.application.md)** object, you need to complete the following three steps:

1. Declare an object variable in a class module to respond to the events.
    
2. Write the specific event procedures.
    
3. Initialize the declared object from another module.
    

## Declare the object variable

Before you can write procedures for the events of the **Application** object, you must create a new class module and declare an object of type **Application** with events. For example, assume that a new class module is created and called EventClassModule. The new class module contains the following code.

```vb
Public WithEvents App As Publisher.Application
```


## Write the event procedures

After the new object has been declared with events, it appears in the **Object** drop-down list box in the class module, and you can write event procedures for the new object. When you select the new object in the **Object** box, the valid events for that object are listed in the **Procedure** drop-down list box. Select an event from the **Procedure** drop-down list box; an empty procedure is added to the class module.

```vb
Private Sub App_DocumentOpen() 
 
End Sub
```


## Initialize the declared object

Before the procedure will run, you must connect the declared object in the class module (App in this example) with the **Application** object. You can do this with the following code from any module.

```vb
Dim X As New EventClassModule 
Sub Register_Event_Handler() 
 Set X.App = Publisher.Application 
End Sub
```

Run the Register_Event_Handler procedure. After running the procedure, the App object in the class module points to the Microsoft Publisher **Application** object, and the event procedures in the class module will run when the events occur.

> [!NOTE] 
> For information about creating event procedures for the **Document** object, see [Using events with the Document object](using-events-with-the-document-object-publisher.md).



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]