# import pandas as pd
# from IPython.display import display
# from openpyxl import load_workbook

# tabela = pd.read_excel("Migracao.xlsx", sheet_name="Processos  - Preenchimento", header=1)
# display(tabela)
# tabela_data=[]
# for coluna in tabela:
#     tabela_data.append({
#         'Título': coluna['nome']
#     })

import requests

offset = ''
data_task = []

url_base = 'https://app.asana.com/api/1.0/sections/1201422389995890/tasks?opt_fields=&limit=1&{offset}'.format(
    offset=offset)
headers = {
    'Authorization': 'Bearer 2/1202126083545971/1207182145066539:9de56ef74e356640cc82f4a05c7fd7b7',
    'Content-Type': 'application/json'
}


def get_task_data(url, headers):
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()
    else:
        print('Falha na requisição. Código de status:', response.status_code)
        return None


while True:
    data = get_task_data(url_base, headers)
    if not data:
        break

    tasks = data.get('data')
    if tasks:
        for task in tasks:
            gid = task['gid']
            url_task = "https://app.asana.com/api/1.0/tasks/{gid}".format(gid=gid)
            task_data = get_task_data(url_task, headers)
            if not task_data:
                continue

            data_json = task_data.get('data')
            if not data_json:
                continue

            print(data_json['name'])

            url_story = "https://app.asana.com/api/1.0/tasks/{gid}/stories".format(gid=gid)
            story_data = get_task_data(url_story, headers)
            if not story_data:
                continue

            data_story = story_data.get('data')

            for story in data_story:
                if story['type'] == 'comment' and 'Processo: ' in story['text']:
                    processonumero = story['text'].split('Processo: ')[1].split()[0]
                    break
            else:
                processonumero = None

            envolved = objective = None
            if 'tags' in data_json and len(data_json['tags']) > 2 and data_json['tags'][2] is not None and 'name' in data_json['tags'][2]:
                envolved = data_json['tags'][2]['name']
            if 'tags' in data_json and len(data_json['tags']) > 2 and data_json['tags'][1] is not None and 'name' in data_json['tags'][1]:
                objective = data_json['tags'][1]['name'] 

            data_task.append({'name': data_json['name'], 'numero': processonumero, 'envolvido': envolved,
                              'objeto': objective})
            print(data_task)

        if data.get('next_page'):
            offset = 'offset=' + data['next_page']['offset']
            url_base = 'https://app.asana.com/api/1.0/sections/1201422389995890/tasks?opt_fields=&limit=1&{offset}'.format(
                offset=offset)
        else:
            break

    else:
        print('Não há mais tarefas.')
        break

# print(data_task)
# else:
#     print('Falha na requisição. Código de status:', response.status_code)