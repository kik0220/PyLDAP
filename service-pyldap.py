import os, yaml
import gradio as gr
from ldap3 import Server, Connection, ALL, SUBTREE
import win32serviceutil
import win32service
import win32event
import servicemanager
import sys

config = {}
script_dir = os.path.dirname(os.path.abspath(__file__))
config_path = script_dir + "/conf.yml"
if os.path.exists(config_path):
    with open(config_path, 'r') as stream:
        config = yaml.safe_load(stream)
else:
    exit(1)

class PyLDAP(win32serviceutil.ServiceFramework):
    _svc_name_ = 'PyLDAP'
    _svc_display_name_ = 'PyLDAP'
    server = Server(config["host"], get_info=ALL)
    conn = Connection(server, user=config["user"], password=config["password"])

    def get_result(filter,self):
        display_result = ""
        if not self.conn.bind():
            display_result ='Could not bind to server.'
            return display_result
        self.conn.search(config["baseDN"], filter, search_scope=SUBTREE, attributes=['*'])
        for entry in self.conn.entries:
            display_result = entry
        self.conn.unbind()
        return display_result

    def get_user(input_user,self):
        global display_result
        display_result = self.get_result(f'(&(objectclass=person)(sAMAccountName={input_user}))')
        return display_result

    def get_group(input_group,self):
        global display_result
        display_result = self.get_result(f'(&(objectclass=group)(CN={input_group}))')
        return display_result

    def get_computer(input_computer,self):
        global display_result
        display_result = self.get_result(f'(&(objectclass=computer)(CN={input_computer}))')
        return display_result

    def get_userlist(self):
        display_result = ""
        if not self.conn.bind():
            display_result ='Could not bind to server.'
            return display_result
        self.conn.search('CN=Users' + config["baseDN"], '(objectclass=person)', attributes=['sAMAccountName','cn','mail','scriptPath','description'])
        for entry in self.conn.entries:
            display_result += f'{entry.sAMAccountName.value}\t{entry.cn.value}\t{entry.mail.value}\t{entry.scriptPath.value}\t{entry.description.value}\n'
        self.conn.unbind()
        return display_result

    def get_grouplist(self):
        display_result = ""
        if not self.conn.bind():
            display_result ='Could not bind to server.'
            return display_result
        self.conn.search('CN=Users' + config["baseDN"], '(objectclass=group)', attributes=['name','description'])
        for entry in self.conn.entries:
            display_result += f'{entry.name.value}\t{entry.description.value}\n'
        self.conn.unbind()
        return display_result

    def get_computerlist(self):
        display_result = ""
        if not self.conn.bind():
            display_result ='Could not bind to server.'
            return display_result
        page_size = 1000
        cookie = None
        entries = []
        while True:
            self.conn.search(search_base=config["baseDN"],search_filter='(objectclass=computer)',search_scope=SUBTREE,attributes=['cn','whenChanged','operatingSystem','description'],paged_size=page_size,paged_cookie=cookie)
            entries.extend(self.conn.entries)
            cookie = self.conn.result['controls']['1.2.840.113556.1.4.319']['value']['cookie']
            if not cookie:
                break
        for entry in entries:
            display_result += f'{entry.cn.value}\t{entry.whenChanged.value}\t{entry.operatingSystem.value}\t{entry.description.value}\n'
        self.conn.unbind()
        return display_result

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        # ここにGradioやその他のコードを書く
        with gr.Blocks(title="PyLDAP") as self.demo:
            with gr.Row():
                with gr.Tab("ユーザー"):
                    input_user = gr.Textbox(placeholder="Enterで確定",show_label=False)
                with gr.Tab("ユーザー一覧"):
                    button_userlist = gr.Button("取得",variant="primary")
                with gr.Tab("グループ"):
                    input_group = gr.Textbox(placeholder="Enterで確定",show_label=False)
                with gr.Tab("グループ一覧"):
                    button_grouplist = gr.Button("取得",variant="primary")
                with gr.Tab("コンピューター"):
                    input_computer = gr.Textbox(placeholder="Enterで確定",show_label=False)
                with gr.Tab("コンピューター一覧"):
                    button_computerlist = gr.Button("取得",variant="primary")
            with gr.Row():
                display_result = gr.Textbox(label="結果", max_lines=50)

            input_user.submit(self.get_user,[self,input_user],display_result)
            input_group.submit(self.get_group,[self,input_group],display_result)
            input_computer.submit(self.get_computer,[self,input_computer],display_result)
            button_userlist.click(self.get_userlist,self,outputs=display_result)
            button_grouplist.click(self.get_grouplist,self,outputs=display_result)
            button_computerlist.click(self.get_computerlist,self,outputs=display_result)
        self.main()

    def main(self):
        self.demo.launch(share=False,show_api=False,server_port=7860,server_name='0.0.0.0')

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(PyLDAP)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(PyLDAP)

