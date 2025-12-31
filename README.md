# Outlook
Outlook COM Interop

```Javascript
function main() {
	
	var folders = json.parse(outlook.list_folders2("max.muster@example.com", ""));
	var folder;
	foreach(folder in folders) {		
		println(folder["Name"]);
	}

	var emails = json.parse(outlook.list_mails("max.muster@example.com", "Posteingang", 0, 10, false, 1000));
	var mail;
	foreach(mail in emails) {
		println(mail["Subject"]);
	}

	var calendar = json.parse(outlook.search_calendar2("max.muster@example.com", "Kalender", "2025-10-28T14:35:12", "2026-03-28T23:35:12", 10));
	var event;
	foreach(event in calendar) {
		println(event["Subject"]);
	}

	var found = json.parse(outlook.search_mails2("max.muster@example.com", "", "", "GIS", "", "", false, "2024-10-28T14:35:12", "2025-12-28T23:35:12", 100, false, 1000));
	var m; 
	foreach(m in found) {
		println(m["Subject"]);
	}
	
}
```
