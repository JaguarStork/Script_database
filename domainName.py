import pandas as pd
import requests

excel_data = pd.read_excel('/Applications/MAMP/htdocs/jaguarStork_database/Script_database/BaseEnCours.xlsx', sheet_name='Feuil1')
# Lecture de fichier excel ainsi que la feuille

DName = excel_data['Denomination']
# Récuparation de la colonne Dénomination

df = pd.DataFrame(DName)
# Convertion en Dataframe

Link = []
# Création d'un tableau link

for index, row in df.head(1).iterrows():
    # Itération dans le dataframe DF ( head = le nombre d'itération max) pour chaque ligne
    
    URLname = row['Denomination']
    # Récupération des noms d'entreprises
    
    url = "https://api.phantombuster.com/api/v1/agent/5375526206492644/launch?output=result-object&argument=%7B%20%20%20%20%20%20%20%20%20%22ignoreList%22%3A%20%22wikipedia.org%5Cnlinkedin.com%5Cnfacebook.com%5Cngoogle.com%5Cnfindthecompany.com%22%2C%20%20%20%20%20%20%20%20%20%22market%22%3A%20%22fr-FR%22%2C%20%20%20%20%20%20%20%20%20%22numberOfLinesPerLaunch%22%3A%20100%2C%20%20%20%20%20%20%20%20%20%22spreadsheetUrl%22%3A%20%22 " + URLname + " %22%20%20%20%20%20%7D"
    # Création de l'url avec le nom de l'entreprise comme argument en JSON
    

    headers = {
        "accept": "application/json",
        "X-Phantombuster-Key-1": "7neQnNwKSbfEYlGbhj797o52VaT4AQqVkAxA5WUWLCo"
    }
    # Header avec la clé de l'API phantom Buster

    response = requests.post(url, headers=headers)
    # Envoie de la requete et récupération dans une variable de la réponse
    
    
    testResult = response.json()['data']['resultObject']
    # récupère le contenu de data -> resultObject de la réponse 
    
    if testResult is not None:
        nomDomaine = response.json()['data']['resultObject'][0]["link"]
        # récupère le contenu Link de la cellule 0 du tableau resultObject 
        Link.append((URLname, nomDomaine))
        # Et puis l'insère dans le tableau link avec le nom de l'entreprise
        
    else:
        Link.append((URLname, 'NaN'))  
        #  sinon insère le nom de l'entreprise et NaN
        
    
# print(Link)
Link_df = pd.DataFrame(Link, columns=['Entreprise', 'Site web'])
# Création d'un dataframe avec le tableau DF et le nom des colonnes 

# print(Link_df)
    
with pd.ExcelWriter('/Applications/MAMP/htdocs/jaguarStork_database/Script_database/BaseEnCours.xlsx',mode='a',if_sheet_exists='replace') as writer:  Link_df.to_excel(writer, sheet_name='Website', index=False)
# Création du excel a partir du dataframe link_df