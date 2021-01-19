from flask import Flask
from flask_restful import Resource, Api, reqparse
import json
import socket
import wmi
import win32process
from win32gui import GetForegroundWindow
import pythoncom
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import threading
import time
import os
import psutil


def check_settings():
    if not os.path.exists('server_data.json'):
        with open('server_data.json', 'w') as f:
            json.dump({"exe": None, "vol": None, "refresh": False}, f, indent=2)


def get_host_name():
    settings_json = 'settings.json'
    temp_host = ''
    try:
        with open(settings_json, 'r') as f:
            temp_host = json.load(f)['host']
    except FileNotFoundError:
        temp_host = socket.gethostbyname(socket.gethostname())
        with open(settings_json, 'w') as f:
            json.dump({'host': temp_host}, f, indent=2)
    return temp_host


def get_host_name_new():
    get_host = socket.gethostbyname(socket.gethostname())
    settings_json = 'settings.json'
    try:
        with open(settings_json, 'r') as f:
            user_host = json.load(f)['host']
    except FileNotFoundError:
        with open(settings_json, 'w') as f:
            json.dump({'host': get_host}, f, indent=2)
    else:
        if user_host != get_host:
            return user_host, user_host
    return '0.0.0.0', get_host


def get_vol(process):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == process:
            return volume.GetMasterVolume()


def change_vol(value, process):
    cur_val = float(value)/100
    cur_val = min(1.0, max(0.0, cur_val))
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and session.Process.name() == process:
            volume.SetMasterVolume(cur_val, None)


def get_current_process():
    pythoncom.CoInitialize()
    c = wmi.WMI()
    exe = None
    try:
        _, pid = win32process.GetWindowThreadProcessId(GetForegroundWindow())
        for p in c.query('SELECT Name FROM Win32_Process WHERE ProcessId = %s' % str(pid)):
            exe = p.Name
            return exe
    except:
        return None


def get_process_and_vol():
    exe = None
    volume = None

    exe = get_current_process()
    try:
        volume = int(get_vol(exe) * 100)
    except TypeError:
        volume = None

    return exe, volume


def save_data(data):
    with open('server_data.json', 'w') as f:
        json.dump(data, f)


def load_data():
    data = None
    with open('server_data.json', 'r') as f:
        data = json.load(f)
    return data


class Main(Resource):
    def get(self):
        data = load_data()
        data['exe'], data['vol'] = get_process_and_vol()
        save_data(data)
        return data, 200

    def put(self):
        parser = reqparse.RequestParser()
        parser.add_argument('vol', required=True)
        args = parser.parse_args()
        print(args)

        data = load_data()

        try:
            new_vol = int(args['vol'])
        except ValueError:
            return {'error': 'type should be an integer out of 100'}, 400
        else:
            change_vol(new_vol, data['exe'])
            data['vol'] = new_vol

        save_data(data)

        return data, 200


def boot_function(display_host):
    time.sleep(1)
    print(f'''\n\tMyVol2 IP: {display_host}\n\tTo change the IP (If using a VPN or Adaptor) edit settings.json''')


def get_adaptor():
    adaptors_addr = psutil.net_if_addrs()
    adaptor_stat = psutil.net_if_stats()
    open_adators = 0

    counter = 0
    ip_list = []
    for adaptor in adaptors_addr:
        adaptor_name = adaptor
        if getattr(adaptor_stat[adaptor], 'isup') is False:
            continue
        for ip in adaptors_addr[adaptor]:
            if getattr(ip, 'family').name == 'AF_INET':
                address = getattr(ip, 'address')
                print(f'[{counter}] {adaptor_name}: {address}')
                ip_list.append(address)
                counter += 1

    if len(ip_list) == 1:
        return ip_list[0]

    number = input(f'Pick adaptor [0-{counter-1}]: ')
    while True:
        try:
            number = int(number)
            print('')
            return ip_list[number]
        except ValueError:
            number = input(f'Insert a number value [0-{counter-1}]: ')
        except IndexError:
            number = input(f'Insert a number between [0-{counter - 1}]: ')


app = Flask(__name__)
api = Api(app)

api.add_resource(Main, '/main')


if __name__ == '__main__':
    display_host = get_adaptor()
    flask_host = display_host
    check_settings()
    # flask_host, display_host = get_host_name_new()
    boot_thread = threading.Thread(target=boot_function, args=[display_host])
    boot_thread.start()
    app.run(host=flask_host, port=5010)
