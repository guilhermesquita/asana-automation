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

offset =  ''
data_task=[]

url_base = 'https://app.asana.com/api/1.0/sections/1201422389995890/tasks?opt_fields=&limit=2&{offset}'.format(offset=offset)
headers = {
    'Authorization': 'Bearer 2/1202126083545971/1207182145066539:9de56ef74e356640cc82f4a05c7fd7b7',
    'Content-Type': 'application/json'
}


response = requests.get(url_base, headers=headers)
data = response.json()

while True:
    # print(data['next_page']['offset'])
    url = 'https://app.asana.com/api/1.0/sections/1201422389995890/tasks?opt_fields=&limit=100&{offset}'.format(offset=offset)
    response = requests.get(url, headers=headers)
    data = response.json()
    
    if data['next_page']['offset']:
        offset = 'offset=' + data['next_page']['offset']
        
    tasks = data.get('data')
    if tasks:
        for task in tasks:
            gid = task['gid']
            url_task = "https://app.asana.com/api/1.0/tasks/{gid}".format(gid = gid)
            response_task = requests.get(url_task, headers=headers)

            if response_task.status_code == 200:
                jsons = response_task.json()
                data_json = jsons.get('data')
                print(data_json['name'])
                
                url_story = "https://app.asana.com/api/1.0/tasks/{gid}/stories".format(gid = gid)
                response_story = requests.get(url_story, headers=headers)
                json_story = response_story.json()
                data_story = json_story.get('data')

                for story in data_story:
                    if(story['type'] == 'comment'):
                        if 'Processo: ' in story['text']:
                            processonumero = story['text'].split('Processo: ')[1].split()[0]
                            # print(processonumero)
            if 'tags' in data_json and len(data_json['tags']) > 2 and data_json['tags'][2] is not None and 'name' in data_json['tags'][2]:
                envolved = data_json['tags'][2]['name']
            if 'tags' in data_json and len(data_json['tags']) > 2 and data_json['tags'][1] is not None and 'name' in data_json['tags'][1]:
                objective = data_json['tags'][1]['name'] 
            data_task.append({'name': data_json['name'], 'numero': processonumero, 'envolvido': envolved, 'objeto': objective})

        print(data_task)
    
    else:
        print('Falha na requisição. Código de status:', response.status_code)
# else:
#     print('Falha na requisição. Código de status:', response.status_code)