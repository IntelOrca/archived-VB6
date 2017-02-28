=============================================
Populous Palette Editor Readme.txt
=============================================

Welcome to the Populous Palette Editor, which can amend the palette of 256
colours used to render Populous graphics. You can use it to create interesting
new colour effects for your own levels, or to change the appearance of the
supplied levels. 

This file can be read more easily viewed with "word wrap" enabled.

---------------------------------------------
1.)  Table of Contents
---------------------------------------------

  1.) Table of Contents
  2.) Installation/ Uninstallation
  3.) System Requirements
  4.) Package Contents
  5.) Using the Populous Palette Editor
  6.) Version History & Known Bugs
  7.) Feedback
  8.) Credits

---------------------------------------------
2.)  Installation/ Uninstallation
---------------------------------------------

To install the Populous Palettte Editor, run the file, "setup.exe" by double 
clicking on it.  Then follow the instructions.

If you later decide that you do not wish to keep the Populous Palette Editor you 
can uninstall it by going to Start\Settings\Control Panel\Add Remove 
Programs. Find Populous Palette Editor and click Add/Remove (change/remove in windows XP and 2000).
Then follow the on-screen instructions.

---------------------------------------------
3.)  System Requirements
---------------------------------------------

Windows 95 or newer.
Any Monitor and video card that is capable of handling a resolution of 1024x768 or greater.
It uses less than 6Mb of RAM and 300Kb of free hard disk space.

---------------------------------------------
4.)  Package Contents
---------------------------------------------

Setup.exe (installer)
Populous Palette Editor (executable)
readme-Help.txt	(this help file)
Descriptions.txt (descriptions of palette entries)

---------------------------------------------
5.)  Using the Populous Palette Editor
---------------------------------------------

---- Before you start

You will need to find out which palette file your level uses. This is coded
in the levl????.hdr file for your map - i.e. the map texture code letter if you use
ALACN's HDR editor.

You can also use the editor on the palette for the flash screen which comes up before the main menu (fepal1.dat in data/fenew), or the main menus (fepal0.dat in data/fenew)

---- Run the program

Click on Populous Palette Editor to run the program. A blank palette will appear with a grid.
Scrolling the grid will show all 256 entries in the palette. Descriptions against the entries
give you an idea of where the colours are used. 

---- Load a palette file

File Open your palette file. This is pal0-?.dat in the Populous data folder, where ? is the map
texture code letter found as described above.

---- The structure of the palette

The Palette is divided up into a number of areas. The descriptions give you an idea what will
be affected if you change the colour. The first half of the palette is used for texturing land
and water (which is used to create textures in the bigf?.dat files), the second half is fixed
for all levels supplied with Populous or Undiscovered Worlds.

Note that since the Populous designers fixed the second half of the palette, these colours are
also used for texturing land, panel displays, menus etc. Consequently if you change them, you
cannot avoid some side effects to land/hut/panel rendering etc.

A good start is to try modifying the tribe colours, which are in groups of 8 adjacent colours
in descending intensity. Another exciting effect is to change index 129 to a bright colour
like pink to see the effect on smoke, spells, map plans etc.    

---- changing the colours

There are several ways of editing a colour. Firstly you can double-click a row in the grid to
select a colour as you would in Paint. This is also available from the Edit menu (Ctrl+E).
secondly you can change the Red/Green/Blue values directly in the Grid (there is no overtype
so use backspace or delete to remove the existing value). Finally you can change the 
Hue/Luminosity/Intensity values of the colour. Keeping the Hue and Luminosity for a group of
colours the same, while changing the Intensity, is a good way of getting shades of the same
colour.

---- Replace

The Replace function (Ctrl+R) is available on the edit menu. This allows you to do block edits by replacing multiple colours in the palette using a wild card search and replace on either Red/Green/Blue or Hue/Saturation/Luminosity. You can restrict the number of rows which will be searched. For example, if you want to replace the blue tribe colours with a different set of shaded colours, you could do an HSL Replace: ?/?/? With: 200/200/? on Rows 216-223. Because you have a wildcard on luminosity, which shades the blue colours from dark to light, the resulting change will give you a range of pinks, shaded from dark to light in a similar way. Take care with Replace as there is no Undo.

---- Other editing features

You can Undo the last change of colour, and copy and paste individual rows. Cut is the same 
as copy, but will set the current row to Black. File New will set the whole grid to Black,
so don't do this accidentally, as Undo will not work. 

---- Edit Tribe Colours

This allows you to easily change a tribe's colour in the palette. Just pick a colour using the selected palettes or using a custom colour. The program will automatically select 7 gradients from that colour and save it onto your palette.

---- Saving your file

Save & SaveAs are available, but make sure that the files are not read-only. Saving a copy
of your palette before you start is a good idea.

---- Editing the Descriptions File

If you undertake your own, more detailed analysis of the palette files, you can update the
descriptions shown in the grid by amending descriptions.txt in the installation directory.
Edit it with care!

---------------------------------------------
6.)  Version History & Known Bugs
---------------------------------------------
Version 0.1 - Pre Release 

There are some minor undesirable features caused by Visual Basic 6.0 controls which I have not managed to eliminate.

1. The colour is still changed if you press Cancel on the Edit Color dialog. If this happens
Pressing Undo should reinstate the original colour.

2. If you Cancel out of SaveAs, your file becomes unnamed and you must use SaveAs again, even
if you want to save it back to its original name.

3. Some slight rounding differences can occur on the Hue/Sat/Lum values and they may not quite
correspond to the values shown in the Edit Color dialog.

Version 0.2 - Pre Release

1. Global Replace function added.

2. Descriptions improved.

3. Help updated.

Version 1.3 - Pre Release

1. Cancel on Edit Colour is fixed.

2. Added Edit Tribe Colours.

---------------------------------------------
7.)  Feedback          ted@brambles.org
---------------------------------------------

Feedback on the program is welcome. Tell me about any bugs, and if you are lucky I
might consider fixing them. Also, if you have improved on the descriptions, why not
send me your descriptions.txt and I will incorporate your changes into future updates.

----------------------------------------------
8.)  Credits
----------------------------------------------
	
Designer 	        TedTycoon 
Programmer 		TedTycoon
Technical Support	ALACN