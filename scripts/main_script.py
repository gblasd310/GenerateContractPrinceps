import subprocess
scripts_paths = (
    "C:/Contratos_Princeps/scripts/script_1.py",
    "C:/Contratos_Princeps/scripts/script_2.py",
    "C:/Contratos_Princeps/scripts/script_3.py",
    "C:/Contratos_Princeps/scripts/script_4.py",
    "C:/Contratos_Princeps/scripts/script_5.py",
    "C:/Contratos_Princeps/scripts/script_6.py"
)
ps = [subprocess.Popen(["python", script]) for script in scripts_paths]
exit_codes = [p.wait() for p in ps]
if not any(exit_codes):
    print("Todos los procesos terminaroin con éxito")
else:
    print("Algunos procesos terminaron de forma inesperada.")