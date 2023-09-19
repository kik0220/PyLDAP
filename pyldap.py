import os, yaml
import gradio as gr
from ldap3 import Server, Connection, ALL, SUBTREE
import subprocess

config = {}
script_dir = os.path.dirname(os.path.abspath(__file__))
config_path = script_dir + "/conf.yml"
if os.path.exists(config_path):
    with open(config_path, 'r') as stream:
        config = yaml.safe_load(stream)
else:
    exit(1)

server = Server(config["host"], get_info=ALL)
conn = Connection(server, user=config["user"], password=config["password"])

def get_result(filter):
    display_result = ""
    if not conn.bind():
        display_result ='Could not bind to server.'
        return display_result
    conn.search(config["baseDN"], filter, search_scope=SUBTREE, attributes=['*'])
    for entry in conn.entries:
        display_result = entry
    conn.unbind()
    return display_result

def get_user(input_user):
    global display_result
    display_result = get_result(f'(&(objectclass=person)(sAMAccountName={input_user}))')
    if display_result == "":
        display_result = get_result(f'(&(objectclass=person)(displayName=*{input_user}*))')
    if display_result == "":
        display_result = get_result(f'(&(objectclass=person)(description={input_user}))')
    if display_result == "":
        display_result = get_result(f'(&(objectclass=person)(userPrincipalName=*{input_user}*))')
    if display_result == "":
        display_result = get_result(f'(&(objectclass=person)(sAMAccountName=*{input_user}*))')
    return display_result

def get_group(input_group):
    global display_result
    display_result = get_result(f'(&(objectclass=group)(CN={input_group}))')
    return display_result

def get_computer(input_computer):
    global display_result
    display_result = get_result(f'(&(objectclass=computer)(CN={input_computer}))')
    return display_result

def get_userlist():
    display_result = ""
    if not conn.bind():
        display_result ='Could not bind to server.'
        return display_result
    conn.search('CN=Users,' + config["baseDN"], '(objectclass=person)', attributes=['sAMAccountName','cn','mail','scriptPath','description'])
    for entry in conn.entries:
        display_result += f'{entry.sAMAccountName.value}\t{entry.cn.value}\t{entry.mail.value}\t{entry.scriptPath.value}\t{entry.description.value}\n'
    conn.unbind()
    return display_result

def get_grouplist():
    display_result = ""
    if not conn.bind():
        display_result ='Could not bind to server.'
        return display_result
    conn.search('CN=Users,' + config["baseDN"], '(objectclass=group)', attributes=['name','description'])
    for entry in conn.entries:
        display_result += f'{entry.name.value}\t{entry.description.value}\n'
    conn.unbind()
    return display_result

def get_computerlist():
    display_result = ""
    if not conn.bind():
        display_result ='Could not bind to server.'
        return display_result
    page_size = 1000
    cookie = None
    entries = []
    while True:
        conn.search(search_base=config["baseDN"],search_filter='(objectclass=computer)',search_scope=SUBTREE,attributes=['cn','whenChanged','operatingSystem','description'],paged_size=page_size,paged_cookie=cookie)
        entries.extend(conn.entries)
        cookie = conn.result['controls']['1.2.840.113556.1.4.319']['value']['cookie']
        if not cookie:
            break
    for entry in entries:
        display_result += f'{entry.cn.value}\t{entry.whenChanged.value}\t{entry.operatingSystem.value}\t{entry.description.value}\n'
    conn.unbind()
    return display_result

def ntfs_acl(input_folder):
    display_result = ""
    command = 'powershell "Get-Acl ' + input_folder + ' | Select-Object Owner | Format-Table -AutoSize -Wrap"'
    result = subprocess.run(command, capture_output=True, text=True, shell=True)
    display_result = result.stdout
    command = 'powershell "Get-Acl ' + input_folder + ' | Select-Object AccessToString | Format-Table -AutoSize -Wrap"'
    result = subprocess.run(command, capture_output=True, text=True, shell=True)
    display_result += result.stdout
    return display_result

with gr.Blocks(title="PyLDAP") as demo:
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
        with gr.Tab("NTFS ACL"):
            input_folder = gr.Textbox(placeholder="Enterで確定",show_label=False)
    with gr.Row():
        display_result = gr.Textbox(label="結果", max_lines=50)

    input_user.submit(get_user,input_user,display_result)
    input_group.submit(get_group,input_group,display_result)
    input_computer.submit(get_computer,input_computer,display_result)
    input_folder.submit(ntfs_acl,input_folder,display_result)
    button_userlist.click(get_userlist,outputs=display_result)
    button_grouplist.click(get_grouplist,outputs=display_result)
    button_computerlist.click(get_computerlist,outputs=display_result)

demo.launch(share=False,show_api=False,server_port=7860,server_name='0.0.0.0')
