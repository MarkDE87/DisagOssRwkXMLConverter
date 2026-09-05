# Raspberry PI Disag VIZ

Dies ist ein Anleitung um einen Raspberry PI als Disag VIZ zu konfigurieren.


## 1.Schritt: 

Einen Order mit dem Name VIZ im Home Directory anlegen. Das Home Directory befindet sich unter /home/pi . Anschließend die VIZ Dateien und Ordner welche sich in der Beamerview.zip befinden in den Ordner VIZ kopieren.

## Schritt 2: 

In diesem Order wird eine neue Shelldatei angelegt. Die Datei besitzt den Name startup.sh

Die Terminal Befehle dazu lauten wie folgt:
```
cd /home/pi/VIZ
sudo nano startViz.sh
```

Der Nano Editor ist geöffnet nun folgendes darin erfassen:

```
#!/bin/bash
echo startViz.sh : Launching VIZ Application
cd VIZ
java -classpath classes:lib/log4j-1.2.8.jar:lib/forms.jar:lib/jcommon-1.0.15.jar:lib/jgoodies-common-1.2.1.jar BeamerView ressources.txt
```


## 3. Die Shell .sh Datei in den Autostart aufnehmen:

```
cd /home/pi

sudo chmod +x startViz.sh

cd ~/.config

cd autostart Hinweis: Wenn der autostart Ordner nicht existiert mittel "mkdir autostart" den Ordner anlegen und anschließend mit "cd autostart" in den Ordner wechseln

sudo nano VNC.desktop
```

Der Nano Editor ist geöffnet nun folgendes darin erfassen:
```
[Desktop Entry]
Type=Application
Exec=/home/pi/startvnc.sh
```

Mittels STRG+X und die Abfrage mit J für Yes bestätigen. Der Nano Editor wird geschlossen und die Datei gespeichert.
