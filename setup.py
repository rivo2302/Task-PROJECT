from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["os","time","datetime"], "includes": ["tkinter","xlsxwriter"] }

# On appelle la fonction setup
setup(
    name = "votre_programme",
    version = "1",
    description = "Votre programme",
    options = {"build_exe": build_exe_options},
    executables = [Executable("fonctions.py")],
)