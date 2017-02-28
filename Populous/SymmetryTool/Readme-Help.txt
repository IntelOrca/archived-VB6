==================================================
Populous Symmetry Tool Version 1.5 Readme-Help.txt
==================================================

Welcome to the Populous Symmetry Tool, which can help level designers with the chore of making multiplayer levels fair, with equal wood, land, height, huts, wildmen etc. It works by creating a mirror image of one half of the map, or mirror image in two directions of one quarter of the map. Both land and objects can be mirrored, and in the case of objects specific to one tribe, they can be modified in colour to create the same number of objects for each tribe.    

This file can be read more easily viewed with "word wrap" enabled.

---------------------------------------------
1.)  Table of Contents
---------------------------------------------

  1.) Table of Contents
  2.) Installation/ Uninstallation
  3.) System Requirements
  4.) Package Contents
  5.) Using the Populous Symmetry Tool
  6.) Version History, Restrictions & Known Bugs
  7.) Feedback
  8.) Credits

---------------------------------------------
2.)  Installation/ Uninstallation
---------------------------------------------

To install the Populous Symmetry Tool, run the file, "install.exe" by double 
clicking on it.  Then follow the instructions.

If you later decide that you do not wish to keep the Populous Symmetry Tool you can uninstall it by going to the program's directory and running "Uninstal".

Then follow the on-screen instructions.

---------------------------------------------
3.)  System Requirements
---------------------------------------------

Windows 95 or newer.
Any Monitor and video card that is capable of handling a resolution of 1024x768 or greater. Minimal disk space is required.

---------------------------------------------
4.)  Package Contents
---------------------------------------------

Zip file containing 

Install.exe (installer)

---------------------------------------------
5.)  Using the Populous Symmetry Tool
---------------------------------------------

---- Before you start

Any Populous editing tools need to be used with care and some knowledge of the the way that Populous works. You can seriously corrupt files using these tools, so make sure that you have adequate backups and you only use the tool on the correct files.

You may also need to use other editors e.g. a header file editor such as Ted Tycoon's Spell Editor to set allies and the number of tribes available on the level.

The program should only be used to edit levl????.dat files found in the Populous Levels directory, or equivalent if you have your own map directories.

---- Run the program

Click on Populous Symmetry Tool to run the program. It is only designed to work on existing correctly formatted levl????.dat files, and is not designed to create one from scratch. 

---- Load a level file

File Open your level file. This is normally levl????.dat in the Populous levels folder, and the program looks for this in the default location. See the following list for the standard level files. You can make some interesting maps by mirroring the supplied levels in one, two, four, or even eight planes.

Standard Level Names:

The Journey Begins	Levl2001	Single player
Night Falls		Levl2002	Single player
Crisis of Faith		Levl2003	Single player
Combined Forces		Levl2004	Single player
Death from Above	Levl2005	Single player
Building Bridges	Levl2006	Single player
Unseen Enemy		Levl2007	Single player
Continental Divide	Levl2008	Single player
Fire in the Mist	Levl2009	Single player
From the Depths		Levl2010	Single player
Treacherous Souls	Levl2011	Single player
An Easy Target		Levl2012	Single player
Aerial Bombardment	Levl2013	Single player
Attacked from all Sides	Levl2014	Single player
Incarcerated		Levl2015	Single player
Bloodlust		Levl2016	Single player
Middle Ground		Levl2017	Single player
Headhunter		Levl2018	Single player
Unlikely Allies		Levl2019	Single player
Archipelago		Levl2020	Single player
Fractured Earth		Levl2021	Single player
Solo			Levl2022	Single player
Inferno			Levl2023	Single player
Journey's End		Levl2024	Single player
The Beginning		Levl2025	Single player
Tutorial		Levl2079	Single player
Aftermath		Levl2060	Single player (UW)
Lava Flow		Levl2056	Single player (UW)
Soul Survivor		Levl2063	Single player (UW)
World Wide Web		Levl2062	Single player (UW)
Human Shield		Levl2057	Single player (UW)
No Man’s Land		Levl2058	Single player (UW)
Protection Racket	Levl2078	Single player (UW)
Prisons			Levl2072	Single player (UW)
Overshadowed		Levl2074	Single player (UW)
Fortress		Levl2061	Single player (UW)
L’Assassine		Levl2059	Single player (UW)
Natural Disaster	Levl2076	Single player (UW)
Hills Devide Us		Levl2080	Multiplayer
Eye Of The Storm	Levl2082	Multiplayer
Two Crabs		Levl2083	Multiplayer
Skirmish		Levl2084	Multiplayer
All Around The World	Levl2094	Multiplayer
Barricade		Levl2085	Multiplayer (UW)
Battlements		Levl2086	Multiplayer (UW)
Cog			Levl2095	Multiplayer (UW)
Sliced Beetle		Levl2096	Multiplayer (UW)
Two Way			Levl2097	Multiplayer (UW)
Multiple Choice		Levl2099	Multiplayer (UW)
Linked Isles		Levl2100	Multiplayer
Skirmish		Levl2109	Multiplayer
Three Way		Levl2110	Multiplayer
Avenging-Angles		Levl2111	Multiplayer
Sandy Castles		Levl2112	Multiplayer
Canyon			Levl2102	Multiplayer (UW)
Angels			Levl2101	Multiplayer (UW)
Three Crabs		Levl2119	Multiplayer (UW)
Two On Two		Levl2120	Multiplayer
Craters			Levl2127	Multiplayer
Dead Sea		Levl2128	Multiplayer
Face Off		Levl2131	Multiplayer
Pressure Point		Levl2133	Multiplayer
Cog			Levl2124	Multiplayer (UW)
Clockwise		Levl2138	Multiplayer (UW)
Walls			Levl2139	Multiplayer (UW)

---- Mirroring Land and Objects

The use of the program is fairly obvious, but there are a few things to note. Firstly when you load a map, it will be displayed in a Populous style world view. If you are used to editing the level files in Hex, you should note that the display is upside down i.e. in a Hex editor the x,y coordinates 0,0 are in bottom left of the map, but in the Populous game view shown by the Symmetry Tool 0,0 is at the top left. 

The tool will mirror from bottom to top and left to right. Therefore if you have both horizontal and vertical axis set, the land and objects in the left hand bottom quadrant will be replicated and all other land/objects lost. In the case of both diagonals, it is the triangle along the left hand vertical edge. If there is no land in these areas, everything will become sea. Of course no damage is done until you save the file, so it is fine to experiment.

Note that carrying out diagonal mirroring can have an odd effect on the South Pole view as the symmetry is truncated at the edges of the map in the North Pole view. This explains the odd effect when you look at the South Pole view with hightlight on, however it is correct for the mirroring which takes place.   

The North Pole view (which corresponds to how the map is laid out in the level file) is used for the mirroring, and the South Pole file is provided to remind you of how it will look from the other side of the world - as it is easy to forget how the land joins up at the edges. You can turn off the South Pole view using Options if you wish.

Symmetry lines are shown as a guide to where the program will apply "reflective" symmetry either side of the line. You can turn these lines off with options too. You can also turn off the highlight for the area which will be mirrored.

You also have the choice whether to mirror land or objects independently. Of course, if you are not careful, you may end up with many of your objects in the sea, and they will drown immediately the level starts.   
 
---- Rotating the map

Rotate map repositions the map data in the level file by a 90 degree clockwise rotation. Repeat this for 180 or 270 degrees. This can help you move an area of the map into the correct position for mirroring. Note that you can rotate the map and objects independently by use of the mirror check boxes.

When you rotate map objects, the orientation of buildings and scenery are also rotated.

---- Moving the map

Move map repositions the map data in the level file. This allows you to move the map into the exact position you want before applying symmetry to it. This is useful of you are using an existing map where the part you want to mirror is not in the bottom left quadrant. You can move the map left, right, up or down in fractions of the map's width i.e. from 1/128th to a half of the map's width. Objects are also moved to correspond with the land position.

---- Changing object colours

Using the "replicate as" option buttons you can decide how any objects specific to one tribe are mirrored. Rotating clockwise the program will take any Blue objects in the first segment and change them to the selected colour in the next segment (Red by default). Any Red objects found will be changed to Yellow be default etc. Then moving on to the next segment (if it is a four way mirror). It will apply the same rules to create the objects in the next segment round. If you set B to B, R to R, Y to Y and G to G, then none of the colours will be changed when mirrored.

---- Trigger Colour

Triggers (e.g. as associated with Stoneheads) are set as colour Blue by default, so v1.3 of the Symmetry Tool does not change the colour of these objects.

---- Orientation change

When you mirror map objects, the orientation of buildings and scenery are also mirrored.

---- Saving your file

Once happy with the changes you can use Save or SaveAs, but make sure that the files are not read-only. Making a copy of all level files before you start is a good idea.

---- Save Image

This feature allows you to save what the North and South box contains. This means it will save every detail in the box, so make sure you turn off symmetry lines etc.

---- Undo & Redo

Undo and Redo will work on any mirror, rotate, load, and move commands. A sequence of consecutive moves are treated as one action by Undo. Redo obviously reverses Undo. You can Undo or Redo up to 10 commands.


---------------------------------------------
6.)  Version History, Restrictions & Known Bugs
---------------------------------------------
Version 1.5 - New features.

1. Rotate introduced

2. Undo & Redo Command added

3. Area to be mirrored is now highlighted

4. Symmetry lines now show at edges of the map and are consistent on the South Pole view.

5. Hitting Cancel on the File Open dialog will not now reload the file.

6. Trigger objects are now replicated to point at their corresponding target locations.

Version 1.4 - Small Updates.

1. Save Image, save a bitmap of what the North and South box contains.

2. Loading files is now easyer because you can open "levl...dat" and ".dat" files now.

Version 1.3 - Enhancements to triggers, object orientation and move map.

1. Move map is easier with separate buttons for each direction.

2. Trigger colour is ignored. This makes sure that triggers remain Blue, otherwise they don't work.

3. The orientation of buildings and scenery is now mirrored eg If a Blue hut was pointing North, you get a Red hut pointing South.

4. Some corrections have been made to this help file. 

Version 1.2 - Addition of Move Map

1. A function to reposition the map in the level file has been added. Note move function will not currently move referenced land locations within objects.

2. The registry updates created by default Microsoft VB code have been removed.

3. The mirror objects check box is now not unchecked or disabled when Show Objects is unchecked.

4. Some corrections have been made to this help file. 

Version 1.1 - First Release 

There are some important points to note replicating objects.

Known Features & Bugs:

1. If you have two shamans in one quarter, and replicate four times, you will get eight shamans. This still works but causes some odd effects. Only one of each colour can cast spells and be reincarnated. 

2. With highlight on and both diagonals symmetry set the South Pole view appears odd, in that the highlighted area appears as a disjointed square and two triangles. This is a feature of the way the Populous map is wrapped into a sphere, but does correctly show the odd effect caused by diagonal mirroring. 

---------------------------------------------
7.)  Feedback          Ted@brambles.org
---------------------------------------------

All use is completely at your own risk, but feedback on the program is welcome. Please tell me about any bugs, and I will consider fixing them. 

----------------------------------------------
8.)  Credits
----------------------------------------------
	
Copyright TedTycoon 2004
