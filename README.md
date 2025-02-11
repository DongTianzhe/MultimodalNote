# MultimodalNote

Note: I did not add many decorations in the game so it looks like a software but please believe it as a game (ゝ∀･)⌒☆

Player can use selection panel (in the view panel or press 'Ctrl+D' to open) to enter details of different modes of moving media like movie or TV series to the main window.

Player can also upload images and music score to the main window. If there is no present music score, player can also enter details like name of notes, music tempo, how many instruments present etc. in the Score(Text) area then generate and copy a LilyPond format that can genreate svg music score using LilyPond software.

Player can add a row at bottom on main window in view menu or press 'Ctrl+B'.

Player can save and open workings on the main window as excel (.xlsx) file. (For unknown reason, the file cannot be saved on Desktop, I will try to fix it later).

The game aim to let player understand how moving media work using KINEIKONIC MODE by letting players analysing these media with several modes or aspects.

By making this game, I have a better understanding of how to use PySide6 and its widgets. In addition, I have learnt how to manage the process of the game so that it can be finished on time.

## Running it from python

1. Open a console and clone the repository:

``` shell
git clone git@github.com:DongTianzhe/MultimodalNote.git
cd MultimodalNote
```

2. Create and activate a new python environment

``` shell
python -m venv .venv 
. .venv/bin/activate # note the . at the beginning of the line
```

3. Install the dependencies

``` shell
pip install -r requirements.txt
```

4. Open the game

``` shell
python main.py
```
