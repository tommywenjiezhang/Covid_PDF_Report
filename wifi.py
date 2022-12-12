
import os


def connect_to_wifi():
    os.system('cmd /c "netsh wlan show networks"')
    router_name = "Resident 4"
    os.system(f'''cmd /c "netsh wlan connect name = "{router_name}""''')

if __name__ == "__main__":
    connect_to_wifi()