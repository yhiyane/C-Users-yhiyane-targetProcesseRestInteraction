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
headers = load_contenu_conf['headers']
querystring = load_contenu_conf['querystring']
url_get_projects = load_contenu_conf['url_get_projects']
url_get_teams = load_contenu_conf['url_get_teams']
url_Select_Project = load_contenu_conf['url_Select_Project']
url_get_Epics = load_contenu_conf['url_get_Epics']
url_get_Features = load_contenu_conf['url_get_Features']
url_get_user_stories = load_contenu_conf['url_get_user_stories']





# Chargement du fichier FormatDataImport qui contient la liste des projects à affecter à une équipe en question
fichier_import_data = open("FormatDataImport.json")
contenu_fichier_import_data = fichier_import_data.read()
fichier_import_data.close()
payload_import_data = contenu_fichier_import_data



# Récupérer la liste des équipes existants
response_get_teams = requests.get(url_get_teams, headers=headers)
list_all_teams = xmltodict.parse(response_get_teams.text)
list_all_teams_json = json.dumps(list_all_teams)
list_all_teams_json_load = json.loads(list_all_teams_json)
list_import_data_teams_load = json.loads(contenu_fichier_import_data)

print("\033[35--------------la liste des équipes existantes -------------------\033[0m")

listTeamsExistant = []
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
for key in list_all_projects_json_load['Projects']['Project']:
    listProjetsExistant.append(key['@Name'])



print("--------------------------------")
print("\033[35m la liste des projets à affectés:\033[0m")
print("--------------------------------")

for item in list_import_data_load:
    print(str(int(item['id']))+"==>"+item['teamProjects'][0]['project']['name'])




# Vérification de l'éxistance d'un projet
for item in list_import_data_load:
   if item['teamProjects'][0]['project']['name'] in listProjetsExistant:
                a = 1
                # print("le projet " + item['teamProjects'][0]['project']['name'] + " existe déjà il faut l'affecté")
   else:
                # si le projet n'existe pas on le crée

                newproject = {
                                "Name": item['teamProjects'][0]['project']['name'],
                                "EntityState": {"Id": 125}
                             }

                if(item['teamProjects'][0]['project']['name'] != ""):
                        response_Insert_Project = requests.post(url_get_projects, data=json.dumps(newproject), headers=headers)
                        ObjectJsonNewProjectsInserted = json.dumps(response_Insert_Project.text)
                else:
                    print("\033[31m le nom du projet est vide!! \033[0m")

# la récupération des id des projets créés

Response_projects = []
hash_name_id_project = []
for item in list_import_data_load:
    if(item['teamProjects'][0]['project']['name'] != ""):
        l = 0
        querybyname = "'" + item['teamProjects'][0]['project']['name'] + "'"
        select_url = url_Select_Project + querybyname
        getObjectsById = requests.get(select_url, headers=headers)
        selected_project = xmltodict.parse(getObjectsById.text)
        selected_project_json = json.dumps(selected_project)
        selected_project_json_load = json.loads(selected_project_json)
        Response_projects.append(selected_project_json_load['Projects']['Project']['@Id'])
        name_id_project = {}
        name_id_project["Name"] = item['teamProjects'][0]['project']['name']
        name_id_project["Id"] = selected_project_json_load['Projects']['Project']['@Id']
        if(len(hash_name_id_project)==0):
            hash_name_id_project.append(name_id_project)
        else:
            for p in hash_name_id_project:
                if (p["Name"] != item['teamProjects'][0]['project']['name']):
                    l = l+1
            if(l == len(hash_name_id_project)):
                hash_name_id_project.append(name_id_project)




print("    \033[35m           ==============Résultat de l'affectation==============\033[0m")
# l'affectation des projets
nbr_succes = 0
nbr_failed = 0
j = 0

response_get_epics = requests.get(url=url_get_Epics, headers=headers)
list_all_epics = xmltodict.parse(response_get_epics.text)

#Récupérer la liste des projets entré dans le fichier excel
list_projects_input = []
for item in globalhash:
    list_projects_input.append(item["project"])


projects_not_insert_epic = []
for p in hash_name_id_project:
    for e in list_all_epics['Epics']['Epic']:
        if (e['@Name'] == "Project Casting & EC" and p["Name"] == e['Project']['@Name']):
            print(""+e['@Name']+"  "+p["Name"])
            projects_not_insert_epic.append(e['Project']['@Name'])


list_projects_not_insert_epic = set(projects_not_insert_epic)

# création des epics features et user stories liés aux projets s'il n'existent pas
k = 0
for item in list_projects_not_insert_epic:
    print(item)

for p in hash_name_id_project:
    if p["Name"] not in list_projects_not_insert_epic:
        new_epic1 = {
            "Name": "Project Casting & EC",
            "Project": {"Id": p["Id"]},
            "EntityState": {"Id": 117},

        }
        new_epic2 = {
            "Name": "GOALS",
            "Project": {"Id": p["Id"]},
            "EntityState": {"Id": 117},
        }
        response_Epics2 = requests.request("POST", url_get_Epics, data=json.dumps(new_epic2), headers=headers)
        response_Epics1 = requests.request("POST", url_get_Epics, data=json.dumps(new_epic1), headers=headers)
        epic_dict = xmltodict.parse(response_Epics1.text)
        epic_id = epic_dict['Epic']['@Id']
        new_feature1 = {
            "Name": "Project Casting",
            "EntityState": {"Id": 105},
            "Project": {"Id": p["Id"]},
            "Epic": {"Id": epic_id}

        }
        new_feature2 = {
            "Name": "Externel Costs",
            "EntityState": {"Id": 105},
            "Project": {"Id": p["Id"]},
            "Epic": {"Id": epic_id}

        }
        response_Features1 = requests.request("POST", url_get_Features, data=json.dumps(new_feature1),
                                              headers=headers)
        response_Features2 = requests.request("POST", url_get_Features, data=json.dumps(new_feature2),
                                              headers=headers)
        feature_dict = xmltodict.parse(response_Features1.text)
        feature_id = feature_dict['Feature']['@Id']
        new_user_story = {
            "Name": "Developer",
            "Project": {"Id": p["Id"]},
            "Feature": {"Id": feature_id},
            "EntityState": {"Id": 75}
        }
        response_user_stories = requests.request("POST", url_get_user_stories, data=json.dumps(new_user_story),
                                                 headers=headers)



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


            #print(response_Epics.text)
            #print(response_Features.text)
            #print(response_user_stories.text)
            payload_project = json.dumps(team_project_relation)
            response = requests.request("POST", url, data=payload_project, headers=headers, params=querystring)
            if response.status_code == 400:
                nbr_failed = nbr_failed + 1
                if (item["project"] != ""):
                    print("\033[33m Bad request, il se trouve qu'il y a déjà une relation entre l'équipe " + str( int(item["id_Team"])) + " et le projet " + (item["project"]) + "\033[0m")
                else:
                    print("\033[31m le projet à affecté à l'équipe " + str(int(item["id_Team"])) + " est vide merci de bien vérifier le nom du projet \033[0m")

            elif response.status_code == 200:
                if(item["project"] != ""):
                    nbr_succes = nbr_succes + 1
                    print("\033[34m le projet " + (item["project"]) + " est affecté à l'équipe " + str(int(item["id_Team"])) + " avec succès\033[0m")
                else:
                    nbr_failed = nbr_failed + 1
                    print("\033[31m le projet à affecté à l'équipe " + str(int(item["id_Team"])) + " est vide merci de bien vérifier le nom du projet \033[0m")

        else:
            nbr_failed = nbr_failed + 1
            print("\033[31m la team numéro : "+str(int(item["id_Team"])) + " n'existes pas merci de s'assurer que l'identifiant de la team est correcte!!\033[0m")

    except :
        nbr_failed = nbr_failed + 1
        print("\033[31m l'identifiant de la team : "+str(int(item["id_Team"]))+" n'est pas numérique Veuillez s'assurer que l'identifiant de la team est un entier!!\033[0m")
print("-------------------------------------------------Résumé-----------------------------------------------------------")
print("")
print("\033[34m Affectations avec succès ==> \033[0m" + str(nbr_succes))
print("\033[31m Affectations erronées ==> \033[0m" + str(nbr_failed))
print("------------------------------------------------------------------------------------------------------------------")