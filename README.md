# printOutlookAttachments:
VBA Macro to print PDF, Word, and Excel attachments of emails in a particular Outlook folder. If you receive lots of invoices as attachments every day and need to print all of them, this macro may work for you.

# How To Install:
First, create a subfolder of your Outlook Inbox named "TO PRINT" (exclude quotes). Then, create a subfolder of the new folder you just created called "PRINTED".

Next, make sure you have your Developer tab enabled on the ribbon in Outlook. If you're unsure how to accomplish this, use a search engine to look up instructions on "how to enable developer tab in outlook".

Once you have the Developer tab enabled, click on the Developer tab and then click Visual Basic. On the left side of the window, right-click in the blank area underneath the Modules folder. Hover over Insert and click Module. You should now see an object called Module1.

Come back to this GitHub repo, copy the contents of printAttachmentsMacro-Module1.txt, then back to Outlook VBA screen, and paste the contents into the big blank area on the right and click Save. You've just saved the code for Module1. Now do the same thing for Module2, starting with right-clicking and inserting a new module.

After you have both Module 1 and 2 saved with the appropriate code, you'll want to assign the macro "Project1.printCertainAttachments" (without quotes) to a button (or use the macro however you'd like, I just prefer to assign it to a button). If you're unsure how to accomplish this, go back to your search engine and look for a guide on "how to assign a macro to a button in Outlook".

That completes everything you'll need to do in Outlook. Now you'll just need a folder on your desktop named "printAttachmentsMacro" (no quotes). Inside that folder, create another folder named "printMacro", and a text file named "lastPrintMacro.txt".

Finally, go here: http://www.columbia.edu/~em36/pdftoprinter.html and click "download it here" towards the top of the page. Place the PDFtoPrinter.exe file in the new folder you created. Much thanks to Edward Mendelson for developing this program and making it available to all of us!

# How To Use:
Place any emails with PDF, Word, or Excel attachments into the "TO PRINT" folder you created in Outlook. Click the new Quick Access Toolbar button you created. Your printer should start printing, emails should move from "TO PRINT" to "PRINTED", and you should see a results window popup. Note: it is best to leave the popup alone until the printer is done printing. Once it's complete, then click "OK". You can check the log in the new folder on the desktop to see attachments that didn't print from the last job.

Working with Win 10 Pro, using VBA version 7.1 in O365 Outlook
