<div align="center">

## How to make VB EXEs under 5K


</div>

### Description

This article will give you some tips and tools to squeeze those big 16/20/32K executables that VB makes when you only use one or two functions and get them to sizes as low as 1.48K.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ion Alex Ionescu](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ion-alex-ionescu.md)
**Level**          |Intermediate
**User Rating**    |4.4 (48 globes from 11 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ion-alex-ionescu-how-to-make-vb-exes-under-5k__1-45413/archive/master.zip)





### Source Code

<p><b><font size="4">Getting the most out of your EXE files</font></b></p>
<p><font size="3">We all know that VB makes very big EXE files, even if you only
add an empty module. I think the minimum size you can normally get from VB is
12K. Add a form, and that will usually go up to 16K. </font></p>
<p><font size="3">Before starting to use the tools I will mention, there are
some basic tricks you can do. (Note that this assumes you are working on a very
small VB application, not a fully-fledged product). </font></p>
<p><b>Deoptimizing your code for speed</b></p>
<p><font size="3">Firstly, since we are aiming for size, try to
remove code like Variable1 = API Call Variable2 = API call(Variable1)
Variable3 = TrimNull(Variable2) and replace it with a single line of code TrimNull(API
Call(API Call)). Try to remove as many variables as you can from your code.
</font></p>
<p><font size="3">Another trick is that VB saves your project name in the
executable, sometimes more then once, as well as your module/form's name. You
can make the file even smaller by renaming your project and components to only
one character. </font></p>
<p><font size="3">Finally, the last big code optimisation you can do (and this
helps speed too) is to replace aliases by the real API name. Use ShellExecuteA
instead of ShellExecute Alias "ShellExecuteA". Even better, if you are
developping only for your computer or a single OS (XP for example), replace the
API Name by their Ordinal Value. An Ordinal is a number that represents the
location of the API inside a DLL. You can use the "Depends" tool that comes with
VB6 to find the ordinals of each function. Then, just use "ShellExecuteA Lib
"user32" Alias "#xxx" where XXX is the ordinal of ShellExecuteA. This will also
make your code much faster, and much smaller. </font></p>
<p><font size="3">The last trick of course, is to compile in P-Code, as once
again, we don't care about speed. </font></p>
<p><font size="3">However, you will probably notice that even after applying
these optimisations, your file still has the same size. That is because VB uses
a very large alignment (the space between sections within an exectuable) and
your file will be padded with NULLs to reach 12K, no matter what. This is where
it gets a bit more complicated. </font></p>
<p><b>Packers</b></p>
<p><font size="3">Your first choice is to use a packer called FSG (look for FSG
1.33 packer on google). The packer will make your file around 3K. </font></p>
<p><font size="3">Or, you can use CompileController, from this site (Note that
the code was "borrowed" from VBPJ by the author) and specifiy /FILEALIGN:0x200
during the linking process. This will make VB create a pure EXE file, but around
5-6K, instead of a file that will run in memory (as FSG). Save yourself the
trouble, and just use FSG instead. It is much better then UPX by the way. </font>
</p>
<p><b>Deleting ressources</b></p>
<p><font size="3">However, you might still want to squeeze more out of your
application's size. I promised 1-2K, not 3K as FSG probably did. Well, after
compiling your executable, open it with "Resource Hacker" (google it) and delete
all the ressources inside (Version Info and Icons). Don't save it just yet, as
packers cannot compress ressource-empty files. The trick here is to create a
1byte file (with notepad) and save it somewhere. Then, still in Resource Hacker,
go to the Action Menu, and select Add New Ressource. Select the file you just
made, and use type 3, name 1. Your file will now have a 1-byte dummy icon. Save
it. </font></p>
<p><font size="3">The 12K executable might now be 11K, or 10.5, or might still
stay at 12K because of the padding, but don't worry, FSG will see the
difference. Use it on the ressource-empty file, and you should reach 2K. </font>
</p>
<p><font size="3">Of course, once again, this article assumed you are making
some small helper application for your main app. It will have no icon, have
unoptimized coding practices, and, depending on which API technique you used,
will only work on a single OS. However, if that is what you are looking for,
this article will help you squish those big nasty 16K files. As final trick, try
using API code instead of Forms/Controls. Look for Viktor E.'s great code
examples.</font></p>
<p><font size="3">I hope everyone had at least something to learn from this!</font></p>

