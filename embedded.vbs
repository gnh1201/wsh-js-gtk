'---------------------------------------------------------------------------------------------------------
'
' VB script demo to use the GTK-server using STDIN
' Implements embedded GTK
'
' Peter van Eerten August 31 2007.
' Tested with GTK-server 2.1.5
'
' Rename this script to '.vbs' extension and double-click it to start
'---------------------------------------------------------------------------------------------------------

' Always declare all variables
Option Explicit

' This variable must be accessible globally
PUBLIC gtkserver

'---------------------------------------------------------------------------------------------------------

' This subroutine may be called from other scripts as well 
PUBLIC SUB Setup_Gtk
    DIM object, configfile, record, terms, func, process, x
    ' Start GTK-server
    SET process = CREATEOBJECT("Wscript.Shell")
    SET gtkserver = process.EXEC("gtk-server -stdin")
    ' Now define global functions from GTK function names
    SET object = CREATEOBJECT("Scripting.FileSystemObject")
    SET configfile = object.OPENTEXTFILE("gtk-server.cfg", 1, False)
    DO WHILE configfile.ATENDOFLINE <> True
	record = configfile.READLINE
	IF LEFT(record, 13) = "FUNCTION_NAME" THEN
	    terms = SPLIT(MID(record, 17), ",")
	    func = "FUNCTION " & LCASE(terms(0))
	    IF INT(terms(3)) > 0 THEN
		func = func & "("
		FOR x = 1 TO INT(terms(3))
		    func = func & "arg" & x
		    IF x < INT(terms(3)) THEN func = func & ","
		NEXT
		func = func & ")"
	    END IF
	    func = func & vbCrLf & "gtkserver.StdIn.WriteLine " & CHR(34) & LCASE(terms(0))
	    IF INT(terms(3)) > 0 THEN
		FOR x = 1 TO INT(terms(3))
		    func = func & " " & CHR(34) & " & CHR(34) & " & "CSTR(arg" & x & ") & CHR(34)"
		    IF x < INT(terms(3)) THEN func = func & " & " & CHR(34)
		NEXT
	    ELSE
		func = func & CHR(34)
	    END IF
	    func = func & vbCrLF & terms(0) & "=gtkserver.StdOut.ReadLine" & vbCrLf & "END FUNCTION"
	    EXECUTEGLOBAL func
	END IF
    LOOP
    configfile.CLOSE
END SUB

'---------------------------------------------------------------------------------------------------------
'
' This is the demo program
'
'---------------------------------------------------------------------------------------------------------

' Setup GTK stuff
CALL Setup_Gtk()

' Declare variables
DIM txt, win, table, button, entry, text, radio1, radio2, cb

' Define GUI with GTK functions - you can use capitals
GTK_INIT "NULL", "NULL"
win = GTK_WINDOW_NEW(0)
GTK_WINDOW_SET_TITLE win, "Visual Basic Script demo program using STDIN"
GTK_WIDGET_SET_USIZE win, 450, 400
table = GTK_TABLE_NEW (50, 50, 1)
GTK_CONTAINER_ADD win, table
button = GTK_BUTTON_NEW_WITH_LABEL("Exit")
GTK_TABLE_ATTACH_DEFAULTS table, button, 41, 49, 45, 49
entry = GTK_ENTRY_NEW
GTK_TABLE_ATTACH_DEFAULTS table, entry, 1, 40, 45, 49
text = GTK_TEXT_NEW ("NULL", "NULL")
GTK_TABLE_ATTACH_DEFAULTS table, text, 1, 49, 8, 44

' Or use small letters for GTK functions
radio1 = gtk_radio_button_new_with_label_from_widget("NULL", "Yes")
gtk_table_attach_defaults table, radio1, 1, 10, 1, 4
radio2 = gtk_radio_button_new_with_label_from_widget(radio1, "No")
gtk_table_attach_defaults table, radio2, 1, 10, 4, 7
gtk_widget_show_all win
gtk_widget_grab_focus entry

DO
    cb = gtk_server_callback("wait")
    IF cb = entry THEN
	txt = gtk_entry_get_text(entry)
	IF LEN(txt) > 1 THEN gtk_text_insert text, "NULL", "NULL", "NULL", txt & CHR(10), -1
	gtk_editable_delete_text entry, 0, -1
    END IF
LOOP UNTIL cb = button

gtk_server_exit
