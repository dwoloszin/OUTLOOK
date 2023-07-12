import subprocess

def install_requirements():
    try:
        subprocess.check_call(['pip', 'install','--upgrade','-r','requirements.txt'])
    except subprocess.CalledProcessError:
        subprocess.check_call(['pip','install','packege==version'])

if __name__ == '__main__':
    install_requirements()    