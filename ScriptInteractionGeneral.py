import requests
import json
import xmltodict
import xlrd
from collections import OrderedDict
import socket


#Vérification la la connexion internet
try:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.connect(("www.google.com", 80))
except:
    print("\033[31mvotre machine n'est pas connectée à internet, merci de vérifier votre connexion!!\033[0m")
    exit()



# Ouvrez le classeur et sélectionnez la première feuille de calcul
try:
    wb = xlrd.open_workbook('Affectation_Team_Projects.xlsx')
except:
    print("fichier inexistant 'Affectation_Team_Projects.xlsx'");
    exit()
try:
    sh = wb.sheet_by_index(0)
except:
    print("la première feuille de calcul est vide!!")
    exit()

# Liste pour contenir des dictionnaires
team_projects = {}
team_projects["teamProjects"] = []
try:
    entete = sh.row_values(range(0, sh.nrows)[0])
    if (entete[0] == ""):
        exit()
except:
    print("la forme de votre fichier est incorrecte ! entete")
    exit()


# Itérer à travers chaque ligne dans la feuille de calcul et extraire les valeurs en dict
try:
    row_value = sh.row_values(range(1, sh.nrows)[0])
except:
    print("les données erronées!!")
    exit()

playlist = {}
try:
    int(row_value[0])
except:
    print("la forme de votre fichier est incorrecte!!")
    exit()


globalList = []
for rownum in range(1, sh.nrows):
    playlist = {}
    projects = OrderedDict()
    row_values = sh.row_values(rownum)
    try:
        id_numerique=int(row_values[0])
        playlist["id"] = id_numerique
        playlist["teamProjects"] = []
        projectT = {}
        projectT["project"] = {}
        projectT["project"]["name"] = row_values[1]
        projectT["isFullProjectAccess"] = True
        playlist["teamProjects"].append(projectT)
        globalList.append(playlist)
    except:
        print("l'identifiant de l'équipe : "+row_values[0]+" n'est pas numérique")



    globalhash = []
for rownum in range(1, sh.nrows):
    hashTeamProject = {}
    projects = OrderedDict()
    row_values = sh.row_values(rownum)
    hashTeamProject["id_Team"] = row_values[0]
    hashTeamProject["project"] = row_values[1]
    globalhash.append(hashTeamProject)

payload_project = json.dumps(globalList)


# Serialize the list of dicts to JSON
j = json.dumps(globalList)
# Write to file
with open('FormatDataImport.json', 'w') as f:
    f.write(j)

# Chargement du fichier de configuration Conf.json
fichier_configuration = open("Conf.json", "r")
contenu_conf = fichier_configuration.read()
fichier_configuration.close()
load_contenu_conf = json.loads(contenu_conf)
url = load_contenu_conf['url']
querystring = load_contenu_conf['querystring']
url_get_projects = load_contenu_conf['url_get_projects']
url_get_teams = load_contenu_conf['url_get_teams']
url_Select_Project = load_contenu_conf['url_Select_Project']

# Chargement du fichier FormatDataImport qui contient la liste des projects à affecter à une équipe en question
fichier_import_data = open("FormatDataImport.json")
contenu_fichier_import_data = fichier_import_data.read()
fichier_import_data.close()
payload_import_data = contenu_fichier_import_data

headers = {
    'Content-Type': "application/json",
    'Authorization': "Basic YWRtaW46YWRtaW4=",
    'Cache-Control': "no-cache",
    'Postman-Token': "c8b013db-72ae-4363-9be3-40e9eb7435d2"
}

# Récupérer la liste des équipes existants
response_get_teams = requests.get(url_get_teams, headers=headers)
list_all_teams = xmltodict.parse(response_get_teams.text)
list_all_teams_json = json.dumps(list_all_teams)
list_all_teams_json_load = json.loads(list_all_teams_json)
list_import_data_teams_load = json.loads(contenu_fichier_import_data)

listTeamsExistant = []
print("\033[35--------------la liste des équipes existantes -------------------\033[0m")

for key in list_all_teams_json_load['Teams']['Team']:
    print("\033[36Id = \033[0m" + key['@Id'] + "\033[36 | Name = \033[0m" + key['@Name'])
    listTeamsExistant.append(key['@Id'])


print("----------------------------------------------")
print("\033[35mVérification de l'existantce d'une team\033[0m")
print("----------------------------------------------")

for item in list_import_data_teams_load:
    if str(int(item['id'])) in listTeamsExistant:
        print("la team num ==> " + str(item['id']) + " existe déjà il faut l'affecté")
    else:
        print("la team num ==> " + str(item['id']) + " n'existe pas !!!")




# récupérer la liste des projets existants
response_get_projects = requests.get(url_get_projects, headers=headers)
list_all_projects = xmltodict.parse(response_get_projects.text)
list_all_projects_json = json.dumps(list_all_projects)
list_all_projects_json_load = json.loads(list_all_projects_json)
list_import_data_load = json.loads(contenu_fichier_import_data)

listProjetsExistant = []


print("---------------------------------")
print("\033[35m la liste des projets existants:\033[0m")
print("---------------------------------")
for key in list_all_projects_json_load['Projects']['Project']:
    print("Id = " + key['@Id'] + " | Name = " + key['@Name'])
    listProjetsExistant.append(key['@Name'])


print("--------------------------------")
print("\033[35m la liste des projets à affectés:\033[0m")
print("--------------------------------")

for item in list_import_data_load:
    print(str(int(item['id']))+"==>"+item['teamProjects'][0]['project']['name'])


print("\033[35m******************* Synthèse ***********************\033[0m")

# Vérification de l'éxistance d'un projet
for item in list_import_data_load:
   if item['teamProjects'][0]['project']['name'] in listProjetsExistant:
                print("le projet " + item['teamProjects'][0]['project']['name'] + " existe déjà il faut l'affecté")
   else:
                # si le projet n'existe pas on le crée
                print("le projet " +item['teamProjects'][0]['project']['name'] + " n'existe pas il faut le créer puis l'affecter")
                newproject = {
                                "Name": item['teamProjects'][0]['project']['name']
                             }
                response_Insert_Project = requests.post(url_get_projects, data=json.dumps(newproject), headers=headers)
                ObjectJsonNewProjectsInserted = json.dumps(response_Insert_Project.text)

# la récupération des id des projets créés

Response_projects = []
for item in list_import_data_load:
    querybyname = "'" + item['teamProjects'][0]['project']['name'] + "'"
    select_url = url_Select_Project + querybyname
    getObjectsById = requests.get(select_url, headers=headers)
    selected_project = xmltodict.parse(getObjectsById.text)
    selected_project_json = json.dumps(selected_project)
    selected_project_json_load = json.loads(selected_project_json)
    Response_projects.append(selected_project_json_load['Projects']['Project']['@Id'])

print("    \033[35m           ==============Résultat de l'affectation==============\033[0m")
# l'affectation des projets

j = 0
for item in globalhash:
    try:
        int(item["id_Team"])
        if (str(int(item["id_Team"])) in listTeamsExistant):
            team_project_relation = {}
            team_project_relation["id"] = item["id_Team"]
            team_project_relation["teamProjects"] = []
            team_project_relation["teamProjects"].append("projet")
            team_project_relation["teamProjects"][0] = {}
            team_project_relation["teamProjects"][0]["project"] = {}
            team_project_relation["teamProjects"][0]["project"]["id"] = Response_projects[j]
            j = j+1
            team_project_relation["teamProjects"][0]["isFullProjectAccess"] = True
            payload_project = json.dumps(team_project_relation)
            response = requests.request("POST", url, data=payload_project, headers=headers, params=querystring)
            if response.status_code == 400:
                 print("\033[33m Bad request, il se trouve qu'il y a déjà une relation entre l'équipe "+str(int(item["id_Team"]))+ " et le projet "+(item["project"])+"\033[0m")
            elif response.status_code == 200:
                print("\033[34m le projet " + (item["project"]) + " est affecté à l'équipe " + str(int(item["id_Team"])) + " avec succès\033[0m")

        else:
            print("\033[31m la team numéro : "+str(int(item["id_Team"])) + " n'existes pas merci de s'assurer que l'identifiant de la team est correcte!!\033[0m")

    except :print("\033[31m l'identifiant de la team : "+item["id_Team"]+" n'est pas numérique Veuillez s'assurer que l'identifiant de la team est un entier!!\033[0m")
