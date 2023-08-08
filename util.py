# Arquivo com métodos úteis
import yaml

def getCredentials():
    
  # pegando as credenciais das apis
  with open("./credentials/credentials.yml","r") as c:
    try:
      credentials = yaml.safe_load(c)
    except yaml.YAMLError as exc:
      credentials = exc
  
  return credentials
