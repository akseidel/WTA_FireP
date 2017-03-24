# WTA_FireP
![RibbonTab](/FPRibbonTab.PNG)

Revit Add-in in c# : Creates a custom ribbon Tab with discipline related tools. The tools at this writing are for placing specific Revit family|types with some families requiring parameter settings made on the fly.

Included is the 3dAimer example that "aims" a special Revit family to a target. The enclosed Revit project is for demonstrating the 3dAimer.

This repository is provided for sharing and learning purposes. Perhaps someone might provide improvements or education. Perhaps it will help to boost someone further up the steep learning curve needed to create Revit task add-ins. Hopefully it does not show too much of the wrong way.  

Used by the tools in this ribbon are classes intended to provide a Revit family instance placement without the Revit use interface overhead normally required by Revit. The classes are intended to provide a universal mechanism for placing some types of Revit families. This includes tags, which is a task not in this discipline Tab but is in other discipline add-ins. The custom tab employs menu methods not commonly explained, for example a split button sets a family placement mode that is exposed to the functions called by command picks. Other tools use Add-in application settings as a way to persist settings or communicate to code that runs subsequent within a command that provides a task workflow.

What goes on in this add-in is much of the typical tasks required for providing a Tab menu interface involving family placement tasks and for implementing those tasks. This means things like:

* Creating a ribbon tab populated with some controls
  - Tool tips
  - Image file to Button image
  - Communication between controls and commands the controls execute
* Establishing the Family|Type for placement
  - Determine if the correct pairing exists in the current file
  - Automatically discovering and loading the family if it does not exist in the current file but does exist somewhere staring from some set directory
* Providing the Family|Type placement interface
  - In multiple mode or one shot mode
  - With a heads up status/instruction interface form
    - As WPF with independent behavior
      - Sending focus back to the Revit view
  - Returning the family instance placement for further processing after the instance has been placed
  - Managing an escape out from the process
  - Handling correct view type context
* Changing family parameter values

Much of the code is by others. Its mangling and ignorant misuse is my ongoing doing. Much thanks to the professionals like Jeremy Tammik who provided the means directly or by mention one way or another for probably all the code needed.
