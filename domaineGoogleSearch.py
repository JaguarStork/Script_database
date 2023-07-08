import pandas as pd
from googlesearch import search  
import time



#  ============ A MODIFIER ============

excel_data = pd.read_excel('/Applications/MAMP/htdocs/jaguarStork_database/Script_database/BaseEnCours.xlsx', sheet_name='Feuil1')
# Lecture de fichier excel ainsi que la feuille

DName = excel_data['Denomination']
# Récuparation de la colonne Dénomination

max_itération = 350
# A 152 une erreur 302 apparait et nous empeche de faire plus de requete ( c'est pourquoi time.sleep permet de contourner le probleme)

#  ============ A MODIFIER ============

start_time = time.time()
# Garde en mémoire le temps de début pour calculer le temps d'execution

count = 0
# Initialise la variable count pour l'afficheur et la pause

df = pd.DataFrame(DName)
# Convertion en Dataframe

Link = []
# Création d'un tableau link

for index, row in df.head(max_itération).iterrows():
    # Itération dans le dataframe DF ( head = le nombre d'itération max) pour chaque ligne
    
    count = count + 1
    # Incrémentation de count
    my_results_list = [] 
    # Création d'une result_list pour contenir les résultats ( vider a chaque nouvelle ligne ) 
    
    if count % 150 == 0:
        # Code à exécuter toutes les 150 itérations
        print("Exécution du code toutes les 150 itérations")
        time.sleep(1200)
    
    
    # ============ A MODIFIER ============
    name = row['Denomination']
    # Récupération des noms d'entreprises 
    # ============ A MODIFIER ============
    
    query = name + " enterprise "
    # Rajout d'un mot pour préciser la recherche
    
    for i in search(query,        # The query you want to run  
                tld = 'com',  # The top level domain  
                lang = 'en',  # The language  
                num = 10,     # Number of results per page  
                start = 0,    # First result to retrieve  
                stop = 10,  # Last result to retrieve  
                pause = 2.0,  # Lapse between HTTP requests  
                ):  
        my_results_list.append(i)
        # Insertion des réponses trouvées
    
    def filtrer_liste(liste):
        mots_exclus = ["facebook", "pagesjaunes", "wikipedia", "instagram", "linkedin"]
        nouvelle_liste = [element for element in liste if not any(mot in element.lower() for mot in mots_exclus)]
        return nouvelle_liste
    # Définition des mots a exclure
    
    nouvelle_liste = filtrer_liste(my_results_list)
    # filtrer la liste my_results_list et l'attribuer dans nouvelle_liste
    
    if len(nouvelle_liste):
        Link.append((name, nouvelle_liste[0]))
        # Ajout du site trouvé ainsi que du no d'entreprise
    else:
        Link.append((name, 'NaN'))  
        #  sinon insère le nom de l'entreprise et NaN
    
    print(str(count) + " / " + str(max_itération))
    # Interface visuelle pour voir l'avancement
        
    
Link_df = pd.DataFrame(Link, columns=['Entreprise', 'Site web'])
# Création d'un dataframe avec le tableau DF et le nom des colonnes 
    
    # ============ A MODIFIER ============
with pd.ExcelWriter('/Applications/MAMP/htdocs/jaguarStork_database/Script_database/BaseEnCours.xlsx',mode='a',if_sheet_exists='replace') as writer:  Link_df.to_excel(writer, sheet_name='Website', index=False)
# Création du excel a partir du dataframe link_df

    # ============ A MODIFIER ============
    
print("Opération executé avec succès")
# Interface visuelle
print("--- %s seconds ---" % (time.time() - start_time))
# Interface visuelle pour affiche rle temps du programme 