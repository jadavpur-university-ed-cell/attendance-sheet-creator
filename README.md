# Attendance Sheet Creator

This is an automated attendance sheet creator. You need to add the name of the sessions in the `session_titles` list. For each session, an individual excel sheet is created with names of random students from the `names.xlsx` file. Keep updating the file with more names. 

If for certain session, a single department name or year is required in the attendance sheet, then they are handled in the *IF* block of the code. You can input your requirements there. Make sure you add the title of the event in the IF statement with *or*. For all other sessions, they are handled automatically in the *ELSE* section, you don't have to do anything there.
