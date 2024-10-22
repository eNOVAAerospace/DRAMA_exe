from pandas import ExcelWriter, DataFrame, read_excel
import os
import simplekml
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET

def Data_names(Repository_dir):
    files_names = os.listdir(Repository_dir)
    return files_names

def read_display_config(filename):
    config = {}
    i = 1
    with open(filename, 'r') as file:
        for line in file:
            print(i,'\\',line)
            key, value = line.strip().split('=')
            config[key.strip()] = value.strip()
            i+=1
    return config

def write_config(config, filename):
    with open(filename, 'w') as file:
        for key, value in config.items():
            file.write(f"{key}={value}\n")
    print("Configuration écrite avec succès !")

def modif_settings(config):
    print("Quel paramètre souhaitez-vous changer ?")
    num = input('Numéro : ')
    os.system('cls')
    print('Assurez-vous d\'écrire un paramètre valide !')
    if num == '1':
        print('Le chemin du répertoire DRAMA est actuellement : ', config['DRAMA_DIRECTORY'])
        config['DRAMA_DIRECTORY'] = input('Nouveau chemin : ')
        
    elif num == '2':
        print('Le chemin du répertoire dépôt du .kml est actuellement : ', config['KML_DIRECTORY'])
        config['KML_DIRECTORY'] = input('Nouveau chemin : ')
        
    elif num == '3':
        print('Le chemin du répertoire dépôt du .xls est actuellement : ', config['EXCEL_DIRECTORY'])
        
        config['EXCEL_DIRECTORY'] = input('Nouveau chemin : ')
    elif num == '4':
        print('Le nom du projet est actuellement : ', config['PROJECT_NAME'])
        
        config['PROJECT_NAME'] = input('Nouveau nom : ')
    elif num == '5':
        print('Le type de simulation est actuellement : ', config['SIMULATION_TYPE'])
        
        config['SIMULATION_TYPE'] = input('Nouveau type : ')
    else :
        print('Numéro de paramètre incorrect')
    write_config(config, 'config.txt')


def excel_create(config, file_name) :
    dir_DRAMA =  config['DRAMA_DIRECTORY']
    project_name = config['PROJECT_NAME']
    simulation_type = config['SIMULATION_TYPE']
    dir_excel_repo = config['EXCEL_DIRECTORY']
    dir_kml_repo = config['KML_DIRECTORY']
    excel_file = file_name
    kml_file = file_name


    dir_repository = os.path.join(dir_DRAMA, "PROJETS", project_name, simulation_type,'REENTRY','output')

    list_files_names = Data_names(dir_repository)
    debris_names = []
    for names in list_files_names:
        if 'Trajectory' in names:
            debris_names.append(names.split('_Trajectory')[0])
        # else:
            # print("Erreur lors de l'acquisition des données. Etes-vous sûr que les fichiers de trajectoire et de thermalHystory ont été généré")
    
    # Récuperer le ObjectName dans le bon ordre
    object_id = []
    for names in debris_names:
        # Étape 1 : Ouvrir le fichier Trajectoire
            with open(os.path.join(dir_repository,names+'_Trajectory.txt'), 'r') as file:
                # Étape 2 : Lire les données
                raw_data = file.readlines()  # Lire toutes les lignes du fichier et les stocker dans une liste
                data = []
                for line in raw_data:
                    if line[:10] == '# ObjectID':
                        phrase = line[18:]
                        object_id.append(phrase[:-1])
    count = 1
    dir_excel_file = os.path.join(dir_excel_repo,excel_file+'.xlsx')
    with ExcelWriter(dir_excel_file, engine='xlsxwriter') as writer:
        # Création de la page index
        data_index = []
        p = 1

        xml_file =  os.path.join(dir_DRAMA, "PROJETS", project_name, simulation_type,'REENTRY','input','objects.xml')

        # Parser le fichier XML
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Récupérer tous les uniqueID sous children
        lolo = []
        for inclusion in root.findall('inclusion'):
            parent = inclusion.find('parent')
            children = inclusion.find('children')
            if not lolo:
                lolo.append([parent.find('uniqueID').text, ''])
            for child in children:
                lolo.append([child.text, parent.find('uniqueID').text])
        
        
        for l in range(len(lolo)):
            for object in root.findall('object'):
                if (object.find('uniqueID').text==lolo[l][0]):
                    lolo[l].append(object.find('material').text)
                    primitive = object.find('primitive')
                    primitive_child = list(primitive)[0]  # Le premier enfant sous <primitive>
                    lolo[l].append(primitive_child.tag)
                    lolo[l].append(object.find('name').text)

        for names in object_id:            
            for y in lolo:
                if (y[0]==names):
                    data_index.append([y[4],p, y[1], y[3], y[2]])
                    p+=1

        df_index = DataFrame(data_index, columns=['Name','Page','Parent Name','Shape','Material'])
        
        df_index.to_excel(writer, index=False, sheet_name='Index')
         # Récupérer l'objet Workbook et Worksheet
        workbook = writer.book
        worksheet = writer.sheets['Index']
        worksheet.set_column('A:A', 20)
        # Insérer l'image
        image_dir  = os.path.join(dir_repository,'sara.AltitudeVsDownrange.png')
        worksheet.insert_image('G1', image_dir,{'x_scale': 0.5, 'y_scale': 0.5})

        for names in debris_names:
        # Étape 1 : Ouvrir le fichier Trajectoire
            with open(os.path.join(dir_repository,names+'_Trajectory.txt'), 'r') as file:
                # Étape 2 : Lire les données
                raw_data = file.readlines()  # Lire toutes les lignes du fichier et les stocker dans une liste
                data = []
                # Étape 3 : Traiter les données
                for line in raw_data:
                    if line[0] != '#':
                        values = line.split()
                        # Extraire les valeurs de chaque colonne et appliquer les conversions (ex : Km -> m)
                        time = float(values[0])
                        altitude = float(values[1])*1000
                        latitude = float(values[2])
                        longitude = float(values[3])
                        velocity = float(values[4])
                        downrange = float(values[5])
                        drag = float(values[6])
                        lift = float(values[7])
                        side = float(values[8])
                        knudsen = float(values[9])
                        mach = float(values[10])
                        flight_path = float(values[11])
                        heading = float(values[12])
                        data.append([time, altitude, latitude, longitude, velocity, downrange, drag, lift, side, knudsen, mach, flight_path, heading]) 

            # Étape 1 : Ouvrir le fichier Aerothermal
            
            with open(os.path.join(dir_repository,names+'_AeroThermalHistory.txt'), 'r') as file:
                # Étape 2 : Lire les données
                raw_data = file.readlines()  # Lire toutes les lignes du fichier et les stocker dans une liste
                # Étape 3 : Traiter les données
                i=0
                for line in raw_data:
                    if line[0] != '#':
                        values = line.split()
                        # Extraire les valeurs de chaque colonne et appliquer les conversions (ex : Km -> m)
                        # Attention entre les versions 3.1 et 3.5 le contenu des fichier aerothermal a été modifié
                        # il n'y a plus l'altidude. Pour que cela fonctionne pour les deux versions on part de la derniere ligne
                        # time = float(values[0])
                        # altitude = float(values[1])*1000
                        temp = float(values[-8])
                        mass = float(values[-7])
                        think = float(values[-6])
                        convective_heat = float(values[-5])
                        radiative_heat = float(values[-4])
                        oxydation_heat = float(values[-3])
                        integrated_heat = float(values[-2])
                        visibility_factor = float(values[-1])
                        data[i].extend([temp, mass, think, convective_heat, radiative_heat, oxydation_heat, integrated_heat, visibility_factor])
                        i+=1

            df = DataFrame(data, columns=['Time', 'Altitude', 'Lat', 'Long', 'Velocity','Downrange', 'Drag', 'Lift', 'Side', 'Knudsen', 'Mach','Flight Path', 'Heading', 'Temp', 'Mass', 'Thick', 'Convective Heat','Radiative Heat','Oxidation Heat','Integrated Heat','VisibilityFactor'])

            # Écrire les données dans le fichier Excel
            # df.to_excel(writer, index=False, sheet_name=sheet_name)
            df.to_excel(writer, index=False, sheet_name=str(count))
            writer
            count+=1
            
    print("Les données ont été écrites avec succès dans le fichier Excel.")

    df = read_excel(dir_excel_file, sheet_name=None)
    
    sheet_name = list(df.keys())

    # Lire les données de trajectoire à partir du fichier Excel
    kml = simplekml.Kml()


    # Charger le fichier XML
    tree = ET.parse(os.path.join(dir_DRAMA, "PROJETS", project_name, simulation_type,'REENTRY','input','sara.xml'))
    root = tree.getroot()

    # Trouver la balise beginDate
    begin_date_element = root.find('.//beginDate')

    # Récupérer la date
    begin_date = begin_date_element.text
    var = 0
    placemarks = []
    for sheet in sheet_name[1:]:
        df = read_excel(dir_excel_file, sheet_name = sheet)
        folder = kml.newfolder(name="TRACK 3DModel")
        
        # Récupère le nom du débris
        debris_name = debris_names[var].split('.')[1]
        
        var+=1

        # Ajouter les coordonnées et les angles à la trajectoire
        track = folder.newgxtrack(name=debris_name)
        track.style.linestyle.width = 1
        track.altitudemode = simplekml.AltitudeMode.absolute

        
        start_datetime = datetime.fromisoformat(begin_date)
        when = []
        for seconds in df['Time']:
            time_delta = timedelta(seconds=seconds)
            new_datetime = start_datetime+time_delta
            when.append(new_datetime.isoformat())

            if (seconds%60 > 0) and (seconds%60 < 1) and (sheet == '24'):
                lat = df.loc[df['Time'] == seconds, 'Lat'].values[0]
                long = df.loc[df['Time'] == seconds, 'Long'].values[0]
                alt = df.loc[df['Time'] == seconds, 'Altitude'].values[0]
                speed = df.loc[df['Time'] == seconds, 'Mach'].values[0]
                mass = df.loc[df['Time'] == seconds, 'Mass'].values[0]
                temp = df.loc[df['Time'] == seconds, 'Temp'].values[0]
                formatted_time = new_datetime.strftime("%H:%M:%S")
                placemarks.append({
                    'name': formatted_time,
                    'time_begin': new_datetime,
                    'speed_kps': speed,
                    'speed_kph': round(speed*3600),
                    'gmt': new_datetime.strftime('%Y-%m-%d %H:%M:%S'),
                    'Time since start': seconds,
                    'latitude': lat,
                    'longitude': long,
                    'altitude': alt,
                    'mach': speed,
                    'mass': mass,
                    'temp': temp
                })
                

        coord = []
        for index, row in df.iterrows():
            lon, lat, alt, angle = row['Long'], row['Lat'], row['Altitude'],row['Heading']
            coord.append([lon, lat, alt])
            
        track.newwhen(when)
        track.newgxcoord(coord)

    # Ajouter un nouveau dossier pour les événements d'atterrissage
    landing_events_folder = kml.newfolder(name="Landing Events")
    # Créer un style pour les points des info bulles
    met_style = simplekml.Style()
    met_style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png'
    met_style.iconstyle.color = simplekml.Color.hex('ff348fff')  # Couleur en format hexadécimal
    met_style.iconstyle.scale = 0.8
    
    # Affichage des info bulles
    for pm in placemarks:
        placemark = landing_events_folder.newpoint(name=pm['name'])
        placemark.styleurl = '#MET'
        placemark.snippet.maxlines = 1
        placemark.timespan.begin = pm['time_begin']
        # Avec compteur
        # placemark.description = f"""<![CDATA[<head><script type='text/javascript' src='https://www.google.com/jsapi'></script><script type='text/javascript'>google.load('visualization', '1', {{packages:['gauge']}});google.setOnLoadCallback(drawChart);function drawChart() {{var data = new google.visualization.DataTable();data.addColumn('string', 'Label');data.addColumn('number', 'Value');data.addRows(1);data.setValue(0, 0, 'Miles Per Second');data.setValue(0, 1, {pm['speed_mps']});var chart = new google.visualization.Gauge(document.getElementById('speedometer'));var options = {{width: 300,height: 300,max: 6.738821878089181,redFrom: 6.06,redTo: 6.74,yellowFrom: 5.05,yellowTo: 6.06,minorTicks: 5.0,majorTicks: [0,1,2,3,4,5,6,7]}};chart.draw(data, options);}}</script></head><body><h2>Current Space Shuttle Data</h2><h3>Quick Data</h3><div id='speedometer'></div><table><tr><td> </td><td>{pm['altitude']}</td></tr>
        # <tr><td> </td><td>{pm['speed_kph']} Kilometers Per Hour</td></tr></table><hr /><h3>Other Data</h3><table><tr><td><b>Greenwich Mean Time (GMT)</b></td><td> {pm['gmt']}</td></tr><tr><td><b>Mission Elapsed Time (MET)</b></td><td> {pm['met']}</td></tr><tr><td><b>Latitude</b></td><td>{pm['latitude']} degrees</td></tr><tr><td><b>Longitude</b></td><td>{pm['longitude']} degrees</td></tr><tr><td><b>Mach Number</b></td><td>{pm['mach']}</td></tr></table><hr /></body>]]>"""
        
        # Sans compteur
        placemark.description = f"""<![CDATA[<head><script type='text/javascript' src='https://www.google.com/jsapi'></script></head><body><h2>Nom du débris Data</h2>
        <h3>Data</h3><table><tr><td><b>Altitude</b></td><td> {pm['altitude']} km</td></tr><td><b>Speed</b></td><td> {pm['speed_kps']} km/s</td></tr><td><b>Speed</b></td><td> {pm['speed_kph']} km/h</td></tr><td><b>Greenwich Mean Time (GMT)</b></td><td> {pm['gmt']}</td></tr><tr><td><b>Time since start</b></td><td> {pm['Time since start']} s</td></tr><tr><td><b>Latitude</b></td><td>{pm['latitude']} degrees</td></tr><tr><td><b>Longitude</b></td><td>{pm['longitude']} degrees</td></tr><tr><td><b>Mach Number</b></td><td>{pm['mach']}</td></tr><tr><td><b>Mass</b></td><td>{pm['mass']} kg</td><tr><td><b>Temperature</b></td><td>{pm['temp']} K</td></tr></tr></table><hr /></body>]]>"""
        placemark.coords = [tuple(map(float, [pm['longitude'],pm['latitude'],pm['altitude']]))]
        placemark.altitudemode = simplekml.AltitudeMode.absolute
        placemark.style = met_style

    # Enregistrer le fichier .kml
    dir_kml_file = os.path.join(dir_kml_repo,kml_file+'.kml')
    kml.save(dir_kml_file)

    print("Le fichier KML a été créé avec succès.")