============================================================
 Network Change TCP/IP (NC)
 Giorgio Brausi - VBCorner
 http://www.vbcorner.net
 vbcorner@vbcorner.net
============================================================
The NC project intend to help developers which have to change 
TCP/IP parameters protocoll frequently to connect to different 
network server.
It is a paranoia insert 'by hand' this parameters each time 
the user change company (network).

This project will not to be perfect or complete, but only a 
started point that you can to adapt to your preferences.

The source is available and modifiable with no limit!
If you modify, bug-fix or enhanced all or part of this project 
I will grant to inform me about this changes, so I can to 
update the project on my web site.

Thank to Mario Raccagni for your technical support, and Doretto 
Roberto for the enanched support to more network cards.

============================================================
IMPORTANT
The program setup will add a link on Startup folder, thus
when Windows is started, NC will be loaded, also.
Due to the /HIDEONSTARTUP parameter will be used, you can
see the tray icon only. 
Right click to the icon open the NC menu.
============================================================

============================================================
VERSION HISTORY
============================================================
Version 1.3.5 - march 01 2006	
        - NEW
          Add Deutsch language. Thank to Patrik Menne!

        - NEW
          Add the option "Auto-select network card", so if
          your computer contain only ONE network card, the 
          'Select NIC' window will not displayed, and the
          your network card is automatically selected. 
          
        - CHANGE - /HIDEONSTARTUP command parameter:
          The previous /AUTO parameter has been replaced by
          the new /HIDEONSTARTUP.
          Now really the NC window will not show on startup.

        - CHANGE - NC window size
          The size of NC window has been enlarged, to prevent
          more lengthy strings for future languages.

============================================================
Version 1.3.2 - february 08 2006
	- BUGFIX - Switch between Profiles:
	  now, when switch to a different profile, first the 
	  previous settings will reset to 0. After this, the 
	  new settings will be applied.

============================================================
Version 1.3.1 - january 18 2006 (internal release)
	- When a profile is activated, if computer ha ONE 
	  network card only, this card will be automatically 
	  selected on form "Select network card" 

============================================================
Version 1.3.0 - october,10 2005
	Author: Doretto Roberto
	- Add 'frmSelectNIC" form to select your Net card
	  (if you have more Net cards).
	Good work Roberto. Thank!

============================================================
Version 1.2.1 - september, 26 2005
	- some bug fixed
	- some little improvement

============================================================
Version 1.1.0 - June, 22 2005
	- MaxLength of txtIP is now set to 3
	- After you digit three numbers on txtIP, the focus 
	  will move to the next control
	- Same thing is you digit the dot '.'
	- Now you can digit the dot '.' in txtIP controls
	- The 'alternative DNS' is now optional
	- Moved the check for some keys to the KeyUp event
	  of txtIP textboxes.
	- add multi-language support (English + Italiano)

	Autore: Doretto Roberto
	Set automatically the Subnet Mask based on the
	IP address
	
============================================================
Version 1.0.0 - June, 13 2005

	First release
