# PPTxECG
‘Death by PowerPoint’ describes those slide shows that drain all life and energy from a room. A major reason for Death by PowerPoint is bloat - slides written like books, relying on word-count to communicate information.

This tools aims to be an ECG, resuscitating slide-sets by guiding slide creators and editors toward points of bloat, and by generally guiding toward better practice.

# Installing
This is written in Python - so:
1. You’ll need [Python3 to run it](https://www.python.org/downloads/) - depending on OS you may already have this.
2. Clone or download this repository and open in a command line.
   * If you use Python for different things, [consider a virtual environment](https://docs.python.org/3/tutorial/venv.html) to contain/isolate this tool.
3. You'll need a few Python dependencies, the simplest is probably to run:-
    ```pip install -r requirements.txt```
4. Run it.
    ```python3 PPTxECGUI.py```

# Usage
On running the tool, you can enter:
* [optionally] an estimated course duration.
* The path to the:-
    * Specific presentation (i.e. a .pptx file)
    * Folder containing a set of presentations (i.e. one or more .pptx files).

On clicking 'analyse' the tool should analyse the slides given.

Switching to the second tab gives a brief overview of the analysis and allows you to request a spreadsheet. The spreadsheet should appear in the same folder as the presentation(s) and gives a more detailed breakdown - highlighting particular 'pain' points in the set. The first 'worksheet' gives an overview of the set. Subsequent work sheets give detail on specific slidesets.

# Known Issues
Weirdly MacOS includes a *very outdated* tcl-tk that works but makes a lot of things look weird/awful. There are numerous ways around this - I used [Homebrew](https://brew.sh) to work on this.

# Possible development directions
## Summarisation
I've been playing with GenSim, [NLTK](https://www.nltk.org) and [spaCy](https://spacy.io), as well as SaaS services like [AWS Comprehend](https://aws.amazon.com/comprehend/) and Google's [Natural-Language](https://cloud.google.com/natural-language/) to get some kind of auto-summarisation (i.e. a way to cut verbose slidesets to keywords). Have had some promising results, but it depends a lot on the slide topic and writing style.