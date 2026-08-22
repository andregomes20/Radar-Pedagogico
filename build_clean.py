import os
import sys
import subprocess
import shutil

# Diretório do projeto no Desktop
desktop_dir = r"C:\Users\andre\OneDrive\Área de Trabalho\analisador-planilhas-python"
os.chdir(desktop_dir)

print(f"Current working directory: {os.getcwd()}")

# 1. Limpar pastas de build antigas antes de iniciar
print("Cleaning old build directories and spec files...")
shutil.rmtree("build", ignore_errors=True)
shutil.rmtree("dist", ignore_errors=True)
if os.path.exists("RadarPedagogico.spec"):
    os.remove("RadarPedagogico.spec")

# 2. Localizar o caminho do pacote customtkinter
import customtkinter
ctk_path = os.path.dirname(customtkinter.__file__)
print(f"CustomTkinter located at: {ctk_path}")

# 3. Preparar comando do PyInstaller com --clean para ignorar caches
cmd = [
    "pyinstaller",
    "--clean", # Forçar limpeza do cache do PyInstaller
    "--noconfirm",
    "--onefile",
    "--windowed",
    "--name=RadarPedagogico",
    "--version-file=file_version_info.txt",
    "--exclude-module=torch",
    "--exclude-module=torchvision",
    "--exclude-module=scipy",
    "--exclude-module=sqlite3",
    "--exclude-module=jedi",
    "--exclude-module=IPython",
    "--exclude-module=notebook",
    "--exclude-module=sklearn",
    "--exclude-module=scikit-learn",
    "--exclude-module=pygments",
    "--exclude-module=ipykernel",
    "--exclude-module=qtconsole",
    "--exclude-module=zmq",
    "--exclude-module=nbformat",
    "--exclude-module=nbconvert",
    f"--add-data=logo.png;.",
    f"--add-data={ctk_path};customtkinter/",
    "app.py"
]

print("Running PyInstaller command:")
print(" ".join(cmd))

try:
    # Executa o PyInstaller e aguarda a conclusão
    subprocess.check_call(cmd)
    print("BUILD SUCCESSFUL!")
    
    # 4. Copiar o executável gerado na pasta dist para a raiz da pasta do Desktop
    exe_src = os.path.join("dist", "RadarPedagogico.exe")
    exe_dest = os.path.join(desktop_dir, "RadarPedagogico.exe")
    
    if os.path.exists(exe_src):
        # Fechar qualquer processo de RadarPedagogico.exe rodando antes de copiar
        try:
            subprocess.run(["taskkill", "/F", "/IM", "RadarPedagogico.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print("Forced termination of any active RadarPedagogico.exe instances.")
        except Exception:
            pass
            
        src_size = os.path.getsize(exe_src)
        print(f"New built executable size: {src_size} bytes")
        
        copied = False
        try:
            if os.path.exists(exe_dest):
                old_dest = exe_dest + ".old"
                if os.path.exists(old_dest):
                    try:
                        os.remove(old_dest)
                    except Exception:
                        pass
                try:
                    os.rename(exe_dest, old_dest)
                except Exception:
                    pass
            shutil.copy2(exe_src, exe_dest)
            print(f"Executable copied to Desktop project root: {exe_dest}")
            copied = True
            
            # Tentar remover o .old temporário
            if os.path.exists(exe_dest + ".old"):
                try:
                    os.remove(exe_dest + ".old")
                except Exception:
                    pass
        except Exception as copy_err:
            print(f"Erro ao copiar executável: {copy_err}")
            
        if copied:
            # Limpar arquivos temporários build/ dist/ e .spec para deixar a pasta limpa!
            shutil.rmtree("build", ignore_errors=True)
            shutil.rmtree("dist", ignore_errors=True)
            spec_file = "RadarPedagogico.spec"
            if os.path.exists(spec_file):
                os.remove(spec_file)
            print("Temporary build directories cleaned up successfully!")
        else:
            print("Keep build files for manual copying due to write error.")
    else:
        print("Error: Executable was not found in dist folder.")
except Exception as e:
    print("Error during build:", e)
    sys.exit(1)
