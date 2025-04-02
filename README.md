# [FR] Extracteur de notes PowerPoint

## Syntaxe des notes

Dans vos notes de diapositive, ajoutez des lignes de doublage avec cette syntaxe :

```
## Max
Ceci est une ligne de doublage.
Ceci est une autre ligne de doublage.

## Alice
C'est Alice qui parle.
C'est toujours Alice qui parle ici.

Après deux retours à la ligne, les notes sont ignorées par l'outil.
```

## Utilisation de l'outil

1. Téléchargez et installez Python via https://www.python.org/
2. Ouvrez un terminal et exécutez les commandes suivantes :
   ```
   python -m pip install python-pptx
   ```
   ```
   python -m pip install python-docx
   ```
3. Téléchargez le fichier <a href="https://raw.githubusercontent.com/aureliendossantos/powerpoint-notes-extractor/main/extract_script.py" download>`extract_script.py`</a> et placez-le dans le même dossier que vos fichiers PowerPoint.
4. Double-cliquez sur le fichier `extract_script.py`. Si vous y êtes invité, choisissez de l'ouvrir avec Python.

# [EN] PowerPoint Notes Extractor

## Notes syntax

In your slide notes, add voice lines with this syntax:

```
## Max
This is a voice line.
This is another voice line.

## Alice
This is Alice speaking.
Still Alice speaking here.

These are other notes discarded by the tool.
```

## Using the tool

1. Download and install Python at https://www.python.org/
2. Open a terminal and run the following commands:
   ```
   python -m pip install python-pptx
   ```
   ```
   python -m pip install python-docx
   ```
3. Download the <a href="https://raw.githubusercontent.com/aureliendossantos/powerpoint-notes-extractor/main/extract_script.py" download>`extract_script.py`</a> file and put it in the same folder as your PowerPoint files.
4. Double-click on the `extract_script.py` file. If prompted, choose to open it with Python.
