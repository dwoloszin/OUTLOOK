import os
import sys
import shutil


script_dir = os.path.abspath(os.path.dirname(sys.argv[0]) or '.')
destination_directory = os.path.join(script_dir, 'import/Analise/')


source_file_path = r"//internal/FileServer/TBR/Network/NetworkAssurance/RegionalNetworkAssurance-SP/RSQA/10.Organização da rede/_EQUIPE PROJETOS_/Controle Rollout TSSR Geral_V1.xlsx"
#destination_directory = r"C:/Users/f8059678/OneDrive - TIM/Dario/@_PYTHON/OUTLOOK/import/Analise"  # Replace with your desired destination path


def copyFiles():
  if os.path.exists(source_file_path):
      if not os.path.exists(destination_directory):
          os.makedirs(destination_directory)

      destination_file_path = os.path.join(destination_directory, "Controle Rollout TSSR Geral_V1.xlsx")

      shutil.copy(source_file_path, destination_file_path)
      print("File copied successfully.")
  else:
      print("Source file path does not exist.")
