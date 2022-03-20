from functools import reduce
from win32com.client import Dispatch

import os, winshell


testpath = 'E:\\Jonas\\Desktop\\'
CWD = os.getcwd()   


def create_shortcut(name, targetToSave, targetToLink):
    path = os.path.join(targetToSave, name + ".lnk" )

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = targetToLink
    shortcut.WorkingDirectory = CWD
    shortcut.save()

def get_directory_structure(rootdir):
    """
    Creates a nested dictionary that represents the folder structure of rootdir
    """
    dir = {}
    rootdir = rootdir.rstrip(os.sep)
    start = rootdir.rfind(os.sep) + 1
    for path, dirs, files in os.walk(rootdir):
        folders = path[start:].split(os.sep)
        subdir = dict.fromkeys(files)
        parent = reduce(dict.get, folders[:-1], dir)
        parent[folders[-1]] = subdir
    return dir    


create_shortcut("test", testpath, "C:\\Users\\")

tree = get_directory_structure("Globales Lernen")


for themenDir in tree["Globales Lernen"]["Lateinamerika"]:
    print(f"selected {themenDir}")
    for landDir in tree ["Globales Lernen"]:
        if landDir == "Lateinamerika": 
            continue
        print(f"searching in {landDir} for {themenDir}")
        try:
            if tree["Globales Lernen"][landDir][themenDir]:
                targetToSave = f"C:\\Users\\locke\\Coding\\GL-links\\Globales Lernen\\Lateinamerika\\{themenDir}\\"
                targetToLink = f"C:\\Users\\locke\\Coding\\GL-links\\Globales Lernen\\{landDir}\\{themenDir}"
                print(f"{themenDir} in {landDir} vorhanden, \n erstelle shortcut in \n {targetToSave} nach: \n {targetToLink}")
                create_shortcut(f"{themenDir} in {landDir}" , targetToSave, targetToLink)
        except KeyError:
            print(f"{themenDir} ist in {landDir} nicht vorhanden")


