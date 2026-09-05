# Raspberry PI VIZ VNC Connect

Dies ist ein Anleitung um einen zweiten Raspberry PI zu konfigurieren welcher mittels VNC Viewer sich auf einen Haupt Rapsberry PI verbindet worauf das Disag VIZ Programm läuft. Hintergrund für diese Anleitung ist:
* Eine Lösung die auch mit WLAN funktioniert, da das VIZ von Disag nur über ein Ethernet Kabel betrieben werden darf. 
* Einen zweites/drittes VIZ zu realisieren ohne ein zweites VIZ zu betreiben.



## 1. Installieren des VNC viewers:
Über die Kommandozeile folgendes ausführen:

 ```
 sudo apt install realvnc-vnc-viewer -y
```

## 2. Eine viz.vnc Config-Datei im Home Verzeichnis anlegen mit folgendem Inhalt. 

```
ChangeServerDefaultPrinter=0
EnableRemotePrinting=0
EnableToolbar=0
FriendlyName=VIZ Main
FullScreen=1
Uuid=ab1ffd4d-8920-4bd3-9d15-f93da9f33299
```

## 3. Eine startvnc.sh Datei im Home Verzeichnis anlegen mit folgendem Inhalt:

Erklärungen:
- Das Sleep mit den 30 Sekunden ist dazu da, dass gewartet wird bis alle anderen Systeme vorallem das VIZ1 mit PI hochgefahren ist.
- IP-Adresse die hier angegeben ist, ist die IP Adresse des ersten VIZ worauf das VIZ läuft. Diese entsprechend anpassen.

```
#!/bin/bash
echo startvnc.sh : Launching VNC Viewer
sleep 30
vncviewer 192.168.0.105 -FullScreen -AutoReconnect -config /home/pi/viz.vnc
```

## 4. Die Shell .sh Datei in den Autostart aufnehmen:

A: cd /home/pi

B: sudo chmod +x startvnc.sh

C: cd ~/.config

D: cd autostart Hinweis: Wenn der autostart Ordner nicht existiert mittel "mkdir autostart" den Ordner anlegen und anschließend mit "cd autostart" in den Ordner wechseln

E: sudo nano VNC.desktop

F: Der Nano Editor ist geöffnet nun folgendes darin erfassen:

```
[Desktop Entry]
Type=Application
Exec=/home/pi/startvnc.sh
```

G: Mittels STRG+X und die Abfrage mit J für Yes bestätigen. Der Nano Editor wird geschlossen und die Datei gespeichert.
