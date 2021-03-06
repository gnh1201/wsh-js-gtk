'------------------------------------------------------------------------
'
' VB script demo to use the GTK-server using FIFO
'
' Peter van Eerten July 4 2004.
' Revised for GTK-server 1.2 at October 10, 2004
' Revised for GTK-server 1.3 at December 5, 2004
'------------------------------------------------------------------------

Option Explicit

Function Gtk(st)

	'Open the pipe and write
	out.Write st
	'Get GTK-server response
	GTK = pin.ReadLine

End Function

'------------------------------------------------------------------------

'Declare variables
DIM process
DIM tmp, win, table, button, entry, text, radio1, radio2, even
Dim fso, out, pin

'Define GTK-server process
SET process = CreateObject("Wscript.Shell")

'Execute GTK-server
process.Run "gtk-server.exe fifo"

'Wait a little so the GTK-server can initialize
WScript.Sleep 1000

'Define object for PIPES
SET fso = CreateObject("Scripting.FileSystemObject")

'Setup access to named pipes
SET out = fso.CreateTextFile("\\.\pipe\out", 0)
WScript.Sleep 500
SET pin = fso.OpenTextFile("\\.\pipe\in", 1)

'Define GUI
GTK("gtk_init NULL NULL")
win = GTK("gtk_window_new 0")
GTK("gtk_window_set_title " & win & " " & chr(34) & "Visual Basic Script demo program using FIFO" & chr(34))
GTK("gtk_widget_set_usize " & win & " 450 400")
table = GTK("gtk_table_new 50 50 1")
GTK("gtk_container_add " & win & " " & table)
button = GTK("gtk_button_new_with_label Exit")
GTK("gtk_table_attach_defaults " & table & " " & button & " 41 49 45 49")
entry = GTK("gtk_entry_new")
GTK("gtk_table_attach_defaults " & table & " " & entry & " 1 40 45 49")
text = GTK("gtk_text_new NULL NULL")
GTK("gtk_table_attach_defaults " & table & " " & text & " 1 49 8 44")

radio1 = GTK("gtk_radio_button_new_with_label_from_widget NULL Yes")
GTK("gtk_table_attach_defaults " & table & " " & radio1 & " 1 10 1 4")
radio2 = GTK("gtk_radio_button_new_with_label_from_widget " & radio1 & " No")
GTK("gtk_table_attach_defaults " & table & " " & radio2 & " 1 10 4 7")
GTK("gtk_widget_show_all " & win)
GTK("gtk_widget_grab_focus " & entry)

DO
    even = GTK("gtk_server_callback wait")
    IF even = entry THEN
	tmp = GTK("gtk_entry_get_text " & entry)
	GTK("gtk_text_insert " & text & " NULL NULL NULL " & chr(34) & tmp & chr(34) & " -1")
	REM Empty entry field
	GTK("gtk_editable_delete_text " & entry & " 0 -1")
    END IF
LOOP UNTIL even = button

'Exit GTK without waiting for answer
out.Write "gtk_server_exit"

'Close handles to the PIPES
out.Close
pin.Close

' End
