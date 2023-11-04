# Mitgliederlisten Generator

Es handelt sich hierbei um ein internes Tool zum Generieren von Excel-Dateien aus Daten des Vereinsverwaltungsprogramms [ComMusic](https://www.commusic.de/start.html).

## Installation
Die einfachste Methode ist es, [Anaconda](https://www.anaconda.com/) zu installieren und eine neue Python-Umgebung aus der [environment.yml](environment.yml) zu erzeugen.
```bash
conda env create -f environment.yml
```

## Ausführen
Die Python-Umgebung wird mit `conda activate reportGenerator` aktiviert.
Danach kann das Skript [Mitgliederlisten_Generator.py](Mitgliederlisten_Generator.py) ausgeführt werden.
