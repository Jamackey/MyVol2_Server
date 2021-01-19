from flask import Flask
from flask_restful import Resource, Api, reqparse
import json
import socket
import wmi
import win32process
from win32gui import GetForegroundWindow
import pythoncom
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

host = socket.gethostbyname(socket.gethostname())


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


app = Flask(__name__)
api = Api(app)

api.add_resource(Main, '/main')

if __name__ == '__main__':
    app.run(host=host, port=5010)
