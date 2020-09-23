README FILE FOR DNA!
CONTENTS
1.) Introduction
2.) System Requirements
3.) Working
4.) Error Reporting
5.) License
6.) FAQs
7.) Credits

INTRODUCTION
 Viruses have made the computers alarmingly defenseless. With advent of each virus, the whole computer world goes into
 a state of chaos. Viruses can master the whole system and steal your credit cards of other sensitive information.
 Moreover, they can crash your whole system and your company's sensitive information will be gone.
 Once a virus makes way into a system then it can bring the system into its knees. But every virus has a weak point.
 That is, it must first be copied into the victim's computer.
 		We all know that the virus needs some medium to get copied. Nowadays Internet and the pen drives are the
 most frequently used mediums. DNA (DNA is Not an Antivirus.) monitors your removable devices and prevents the transmission
 of the virus into your system. Since, prevention is always better than cure, thus, DNA proves, sometimes, itself even better
 than an Antivirus. It's a complete open source package. Thus, you also get the source code of DNA and yourself master
 the tactics to prevent infection. DNA has some advanced features too like immunization, detection of executable folders
 and probability meter. When you get used to it, you surely will find it an important tool.

SYSTEM REQUIREMENTS
 DNA is completely coded in Visual Basic 6.0. Here are its requirements
 1.) Processor (P3 or above)
 2.) Memory	   (10 MB)
 3.) Operating System (Tested on win XP only but will run if has a Visual Basic Runtime library on win98 or higher)
 4.) Extras
         a.)MSVBVM60.dll (Supplied)
	   b.)Comdlg32.ocx (Supplied)

WORKING
 Each time you start DNA, it checks the whole available drive for Autorun entry. Also it does a surface scan of executable
 Folder of each drive. The first time DNA is run, it asks user to disable autorun(EXTREMELY IMPORTANT)	
	User can access DNA via the system tray. The tray has got following menu.
		1.) Enabled     				
		2.) Check A Drive For Exe Folders.  
		3.) Check each drive for virus
		4.) Immunization
		5.) Options
		6.) About
		7.) Help
		8.) Exit

 ENABLED
	This is a checked menu. If it is checked, the DNA monitor is working and vice versa.
	
 CHECK A DRIVE FOR EXE FOLDERS
	It checks a given drive for the existence of executable folders
 
 CHECK EACH DRIVE FOR VIRUS
 	It checks each drive that is currently attached with the system for existence of autorun entry

 IMMUNIZATION
	It immunizes the selected drive so that it may not be infected with virus when used in another computer.

 OPTIONS
	Opens the option dialog

 ABOUT
	Opens the about dialog

 HELP
	Opens this file
 
 EXIT
	Ends the application.


ERROR REPORTING.	 
 When the program shows some unnatural behavior, then email at DevilLabs@gmail.com with the subject "Error_DNA". Please
 describe when the error occurred (What were you doing when error occurred) also include the APP_LOG.log file with the
 email. (File can found in the same directory where there is the program.)

LICENSE
 This program is released under GNU GPL License

FAQ
 1.) Why should I use DNA?
 	DNA prevents your precious pen drives and system from getting infected.

 2.) But I have an Antivirus?
	Most home version of antiviruses may not detect an infection on the pen drive if it is packed with different 
	packers. Antiviruses are definitely superior than DNA but DNA altogether with AV makes you extra safe.

 3.) I have low system resources, won't DNA hang my system?
	DNA uses a very little resource. It sleeps until any new removable drive is attached and when it finds one, it 
 	becomes active.
 
 4.) Do I have the Visual basic runtime already in my computer?
	Visual Basic Runtime (MSVBVM60.dll) is critical for all Visual Basic application. Hence, there is a higher
 	degree of probability that it may be already installed with other programs. However, if you want to check then
	Goto C:\Windows\system32 (Replace C:\ with the drive in which window is installed) and check if there is
	MSVBVM60.dll. If there is not then you can download from internet or from where you downloaded this app.

 5.) I am getting errors or unnatural behaviors.?
	Please refer the Error reporting section.

 6.) I have some queries regarding DNA.
	Email at DevilLabs@gmail.com with your queries, we will try to reply soon.

CREDITS
	This program is Developed by devil labs. Other multimedia stuffs belongs to their respective authors.

