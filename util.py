import yaml

def getCredentials():
    # Get credentials from the APIs
    with open("./credentials/credentials.yml", "r", encoding="ANSI") as c:
        try:
            credentials = yaml.safe_load(c)
        except yaml.YAMLError as exc:
            credentials = exc
  
    return credentials
