# PowerPoint Notes Extractor

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
3. Download the `extract_script.py` file from this project and put it in the same folder as your PowerPoint files.
4. Double-click on the `extract_script.py` file. If prompted, choose to open it with Python.
