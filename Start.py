from fct_folder import *

def main():
    """GUI (Graphic User Interface) permettant la transcription des simulation DRAMA en fichier .kml et .xls
    """

    print('Programme permettant l\'automatisation de la récupération des données dans DRAMA pour la génération de .kml')
    
    # Acquisition and Display of Parameters 
    config = read_display_config('config.txt')
    
    print('Voulez-vous changer les parametres de génération ?')
    
    see = input('(OUI/NON) : ')
    if (see == 'OUI' or see =='o' or see =='oui' or see =='O'):
        # Parameter Change
        a = 1
        while(a==1):
            modif_settings(config)
            config = read_display_config('config.txt')
            print("Voulez-vous modifier d'autres paramètres ?")
            more = input('(OUI/NON) : ')
            if (more =='NON' or more =='non' or more =='n' or more =='N'):
                see = 'NON'
                a = 0  
    if (see == 'NON' or see =='n' or see =='non' or see =='N'):
        # Get names for output files
        os.system('cls')
        print('Quel nom souhaitez-vous donner aux fichiers .kml et .xls qui vont être crées ?')
        print('Attention aux caractères trop spéciaux !')
        file_name = input('Nom : ')
        
        # Program launch
        os.system('cls')
        excel_create(config, file_name)

    if (see!='NON' and see!='non' and see!='n' and see != 'OUI' and see !='o' and see !='oui' and see !='O' and see !='N'):
        print('Choix Invalide')
    
    input("Appuyez sur Entrée pour quitter...")

main()

