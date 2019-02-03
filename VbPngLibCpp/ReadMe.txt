========================================================================
            Dynamic link library / static link library
                               VBPng
========================================================================

This project is designed to provide the ability to work with PNG images 
in Visual Basic 6.0 using standard controls. You can either load or 
save PNG images.

How to use?

The dll must be loaded before an image is loaded. If it's neccesary to 
avoid dependencies it can be linked statically.
It requires a modern linker because the native linker has the bug and 
it can't statically link a libarary with imports using -OPT:REF option.
So, when this condition is met it can be statically linked by using
the LinkSwitches option like [vbpng.lib -ENTRY:mainCRTStartup].

How does it work?

This library hijacks the OleLoadPictureEx/OleLoadPicture/OleIconToCursor
functions. When a caller calls this one it checks the status and if the 
function failed it attempts to load the image using GDI+. If successful, 
it creates an object which implements all the interfaces like the original 
one excepting IConnectionPointContainer interface. This library keeps 
track of the object references so you shouldn't unload the dll while any 
object is active. To check if the library can be unloaded call the 
CanUnloadNow function. If it returns S_OK the library can be unloaded.

This library uses the instruction length-disassembler by Ms-Rem 
(Ms-Rem@yandex.ru) ICQ 286370715 with some modification.

by The trick, 2019.

