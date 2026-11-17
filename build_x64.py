import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')

project_dir = r"G:\APP LOC DU LIEU\Tổng hợp\VTC_Bao_Cao"
config_dir = os.path.join(project_dir, "config")
main_script = os.path.join(project_dir, "main.py")

cmd = [
    sys.executable,
    "-m",
    "PyInstaller",
    "--noconfirm",
    "--onedir",
    "--windowed",
    "--name=VTC_Bao_Cao",
    f"--add-data={config_dir};config",
    main_script
]

print("Executing PyInstaller build...")
res = subprocess.run(cmd, cwd=project_dir)
if res.returncode == 0:
    print("\nBUILD SUCCESSFUL!")
    print(f"EXE output: {os.path.join(project_dir, 'dist', 'VTC_Bao_Cao', 'VTC_Bao_Cao.exe')}")
else:
    print(f"\nBUILD FAILED with returncode {res.returncode}")
sys.exit(res.returncode)
