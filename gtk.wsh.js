/* 
 * gtk.wsh.js
 *
 * Windows Scripting Host - Javascript (wsh-js) GTK GUI Example
 * 2020-02-09
 * Go Namhyeon <gnh1201@gmail.com>
 * 
 * Requirements
 * * https://www.gtk-server.org/ (GTK-Server)
 * * https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer (GTK-for-Windows-Runtime-Environment-Installer)
 * * https://www.dlldosyaindir.com/dllindir/libffi-6-bit32.rar (libffi-6.dll)
 * 
 * Notes
 * * GTK-Server is not supported GTK3 in Windows, so you have to use GTK2
 *
 */

function CreateObject(objName) {
	return new ActiveXObject(objName);
}

function CHR(ord) {
	return String.fromCharCode(ord);
}

function GTK(st) {
	gtkserver.StdIn.WriteLine(st);
	return gtkserver.StdOut.ReadLine();
}

var process = CreateObject("Wscript.Shell");
var gtkserver = process.Exec("gtk-server -stdin");

function main() {
	var tmp, win, table, button, entry, text, radio1, radio2, even;
	GTK("gtk_init NULL NULL");
	win = GTK("gtk_window_new 0");
	GTK("gtk_window_set_title " + win + " " + CHR(34) + "Visual Basic Script demo program using STDIN" + CHR(34));
	GTK("gtk_widget_set_usize " + win + " 450 400");
	table = GTK("gtk_table_new 50 50 1");
	GTK("gtk_container_add " + win + " " + table );
	button = GTK("gtk_button_new_with_label Exit");
	GTK("gtk_table_attach_defaults " + table + " " + button + " 41 49 45 49");
	entry = GTK("gtk_entry_new");
	GTK("gtk_table_attach_defaults " + table + " " + entry + " 1 40 45 49");
	text = GTK("gtk_text_new NULL NULL");
	GTK("gtk_table_attach_defaults " + table + " " + text + " 1 49 8 44");

	radio1 = GTK("gtk_radio_button_new_with_label_from_widget NULL Yes");
	GTK("gtk_table_attach_defaults " + table + " " + radio1 + " 1 10 1 4");
	radio2 = GTK("gtk_radio_button_new_with_label_from_widget " + radio1 + " No");
	GTK("gtk_table_attach_defaults " + table + " " + radio2 + " 1 10 4 7");
	GTK("gtk_widget_show_all " + win );
	GTK("gtk_widget_grab_focus " + entry );
	
	while(!(even == button)) {
		even = GTK("gtk_server_callback wait");
		if(even == entry) {
			tmp = GTK("gtk_entry_get_text " + entry );
			if(tmp.length > 1)
				GTK("gtk_text_insert " + text + " NULL NULL NULL " + CHR(34) + tmp + CHR(10) + CHR(34) + " -1");
			// Empty entry field
			GTK("gtk_editable_delete_text " & entry & " 0 -1")
		}
	}

	GTK("gtk_server_exit");
}

main();
