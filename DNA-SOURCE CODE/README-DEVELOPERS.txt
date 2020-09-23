DNA is a small utility that I made under the name of the not-registered Devil Labs. I hope this will help you
learn something as well as protect your PC from pesky viruses. There are somethings that still need to be added in DNA.
A few of them are
*DNA and its dependencies need to be protected from any modifications possible by viruses.
*DNA needs a functionality to recognize the drives it scans and add their information in a database so that a
 good quarantine system can be developed.
*And it needs to be optimized so as to consume even less resources than it does.
And lots of other things.
	Though, it has been tested for a long time but you can't be sure when it comes to bugs. So, if you encounter any
bugs, fix it yourself and if you make some good changes to DNA then redistribute it (It's open source).

Finally, if you want to see how this program works then first of all start with the Main form's Load Event and then You can
bp in this part of code in APIStuffs.bas' NewWindowProc procedure.

    If MSG = WM_DEVICECHANGE And lParam <> 0 And status = True Then
    Debug.Print Hex(wParam)		'Place a bp Here to analyze the code.
     If wParam = DBT_DEVICEARRIVAL Then
		a = FindChanges(lParam)

Although the program is being subclassed but bp will be perfectly fine.

This whole project and any part of it's code are all open. Use it in your code without any hesitation (You don't even
need to mention me.)

That's all
Contact me via PSC

Regards
Pravesh

