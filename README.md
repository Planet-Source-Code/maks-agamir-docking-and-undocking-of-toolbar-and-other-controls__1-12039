<div align="center">

## Docking and UnDocking of toolbar and other controls


</div>

### Description

Update

----

This is an improved version of previos DLL. Now it handles child controls of the container control which is being undocked, plus I've fixed the problem with changing focus. Also I've included test project so you can see how it is suppose to work.

Old 

----

This class allows to undock any control on the form and drag it around. Furthermore, during Undocked state events will be still received at the same place. It is an emulation of the Microsofts abilities to undock panels (ex VB Properties and Toolbar)
 
### More Info
 
There is methods which requires parameters to be passed to it. It is handle to the objects which is being undocked, Caption for the undock form and

position.

It is a self contained Dll, it sould not be required to do any programmimng to undock the control.

It can be a little tricky when dealing with buttons controls. One more thing do not put container object within the container object.

For example frame1 and then frame2 inside of

frame1 and then try to undock frame1, because you will have slight glitches with controls on the frame2


<span>             |<span>
---                |---
**Submitted On**   |2000-10-18 13:03:04
**By**             |[Maks Agamir](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/maks-agamir.md)
**Level**          |Advanced
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1075010182000\.zip](https://github.com/Planet-Source-Code/maks-agamir-docking-and-undocking-of-toolbar-and-other-controls__1-12039/archive/master.zip)








