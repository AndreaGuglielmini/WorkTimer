pyinstaller --noconsole --noconfirm .\WorkTimer.py; cp splash.jpeg .\dist\WorkTimer; Copy-Item .\venv311\Lib\site-packages\pandasgui -Destination .\dist\Worktimer\_internal\ -Recurse; Copy-Item .\venv311\Lib\site-packages\qtstylish -Destination .\dist\WorkTimer\_internal\ -Recurse