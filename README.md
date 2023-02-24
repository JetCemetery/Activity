# Activity
Mouse and keyboard mover

I've implemented some basic scaffolding to accept commandline arguments, but this all needs to be tested, and debugged. A future todo is to remove the import of system.drawing and implement the low level win API to get the curosr position, as well as to set it. During testing the setting of curosr position failed for me, need to explore that.

Had to include the key stroke being sent as well, as mouse moving did not do the trick for me on my main work computer.

How to use program:
All the source code is present for inspection, but you can go ahead and copy the built .exe files save it to a computer with .net installed, and then run it.

After initial 5 second delay the program shall move the mouse every 5 seconds. If the program detects a new mouse position, it will auto quit for you as to not get in your way.