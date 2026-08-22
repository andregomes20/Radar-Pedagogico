import os
import sys
import subprocess
import shutil

# Diretório do projeto no Desktop
desktop_dir = r"C:\Users\andre\OneDrive\Área de Trabalho\analisador-planilhas-python"
os.chdir(desktop_dir)

# 1. Limpar pastas de build antigas do instalador antes de iniciar
shutil.rmtree("build_installer", ignore_errors=True)
if os.path.exists("Instalador_RadarPedagogico.spec"):
    os.remove("Instalador_RadarPedagogico.spec")

# 2. Localizar o caminho do pacote customtkinter
import customtkinter
ctk_path = os.path.dirname(customtkinter.__file__)

# 3. Preparar comando do PyInstaller com --clean para ignorar caches
cmd = [
    "pyinstaller",
    "--clean", # Forçar limpeza do cache do PyInstaller
    "--noconfirm",
    "--onefile",
    "--windowed",
    "--name=Instalador_RadarPedagogico",
    "--version-file=file_version_info.txt",
    "installer.py"
]

# Modificar comando se estivermos emulando ou compilando com arquivos embutidos
exe_src = "RadarPedagogico.exe"
logo_src = "logo.png"

if os.path.exists(exe_src):
    cmd.insert(6, f"--add-data={exe_src};.")
if os.path.exists(logo_src):
    cmd.insert(6, f"--add-data={logo_src};.")
    
# Adicionar CustomTkinter
cmd.insert(6, f"--add-data={ctk_path};customtkinter/")

print("Running PyInstaller command for Installer:")
print(" ".join(cmd))

try:
    # Executa o PyInstaller e aguarda a conclusão
    subprocess.check_call(cmd)
    print("INSTALLER BUILD SUCCESSFUL!")
    
    # 4. Copiar o instalador gerado na pasta dist para a raiz da pasta do Desktop
    inst_src = os.path.join("dist", "Instalador_RadarPedagogico.exe")
    inst_dest = os.path.join(desktop_dir, "Instalador_RadarPedagogico.exe")
    
    if os.path.exists(inst_src):
        # Fechar qualquer processo de Instalador_RadarPedagogico.exe rodando antes de copiar
        try:
            subprocess.run(["taskkill", "/F", "/IM", "Instalador_RadarPedagogico.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass
            
        # Workaround para sobrescrever arquivo se estiver bloqueado
        copied = False
        try:
            if os.path.exists(inst_dest):
                old_dest = inst_dest + ".old"
                if os.path.exists(old_dest):
                    try:
                        os.remove(old_dest)
                    except Exception:
                        pass
                try:
                    os.rename(inst_dest, old_dest)
                except Exception:
                    pass
            shutil.copy2(inst_src, inst_dest)
            print(f"Installer copied to Desktop project root: {inst_dest}")
            copied = True
            
            if os.path.exists(inst_dest + ".old"):
                try:
                    os.remove(inst_dest + ".old")
                except Exception:
                    pass
        except Exception as copy_err:
            print(f"Erro ao copiar instalador: {copy_err}")
            
        if copied:
            # Limpar arquivos temporários build/ dist/ e .spec para deixar a pasta limpa!
            shutil.rmtree("build", ignore_errors=True)
            shutil.rmtree("dist", ignore_errors=True)
            spec_file = "Instalador_RadarPedagogico.spec"
            if os.path.exists(spec_file):
                os.remove(spec_file)
            print("Temporary build directories cleaned up successfully!")
        else:
            print("Keep build files for manual copying due to write error.")
    else:
        print("Error: Installer was not found in dist folder.")
except Exception as e:
    print("Error during installer build:", e)
    sys.exit(1)
