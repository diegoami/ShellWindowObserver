TSHELLWINDOWOBSERVER - Version 1.2

Created by Diego Amicabile
Updated by Michel Hibon - 08/27/2001

******************************************************************************************************************************************************************
This component was born out of frustration, since I wasn't able to connect to an existing instance of Internet Explorer AND 
catch the events it fires. Using this component you can, and you can perform actions when a new instance is started or an 
existing instance moves to a new location.There must be a better way to do it, let me know if you have figured it out. I use timers.

See help.htm for documentation.
See license.txt for license agreements.

REQUIREMENTS

THe Internet Explorer Active X Component must have been imported 
(it is by default in Delphi 5)
You may need to replace SHDOCVW with SHDOCVW_TLB

INSTALLING 

Put the uShellWindowObserver.pas file in a package and compile it.  
IMPORTANT : You may have to remove the line {$DEFINE COOBJS} if you are using Delphi 3.

NOTE

The component will be installed in any version of Delphi (3 or higher)
The demo was created using Delphi 5 so the .DFM file may have to be converted.

Diego Amicabile
Email : diegoami@yahoo.it
Web page : http://www.geocities.com/diegoami

******************************************************************************************************************************************************************
Like Diego, I was researched a method to get the handle of an explorer's instance. When you create a new explorere's instance
with CreateProcess API Function, the PROCESS_INFORMATION returns false data. Why ? I don't know ! I have found this excellent
component, and i have decided to integrate it in my program. But, a little bug appears. I correct it, and i add some improvements :

	Count : The number of active instance
	function SendMessageToWindowByNumber(Num: Integer): LRESULT; : Allow you to send a message to choosen instance
	property SendAMessage : Allow you to choose the type of message you want to send at the choosen instance

I design the demo program for test this component with this improvements. Component and demo program was updated with Delphi 6.

Name:     Michel Hibon (alias : Mickey)
E-mail:     mhibon@ifrance.com
Homepage:   http://www.guetali.fr/home/mhibon

**********************************************************************************************************************

VERSION HISTORY

1.0  Created the component for Delphi 3 and Delphi 5 trial version

1.1  Updated also the component for Delphi 5

1.2  Added the improvements from Michel Hibon and the demo from Randy  
	Changed the property of TShellWindow from ID to NUM

TO DO 

  EOleSysException still happens in the development environment. Why?


  
  





