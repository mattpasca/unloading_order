"""
Route planner for the transports of Gorini Piante.
The user gives input through the 'Costi.xlsx' file: names of the companies we have to deliver to 
(which usually has a good matching with 'Acronimo' in the customer database). The program retrieves
detailed address information and decides optimal unloading order.
A mail for marta transport is printed automatically. The route is printed in an html file.

TODO
    - handle errors for unmatched customers
    - handle unloading order in mail for unmatched customers
    - clean up code
"""
import pandas as pd
import Levenshtein as Lev
import pgeocode as pg
import requests
from geopy.geocoders import Nominatim
import time
import folium
import polyline
 
common_terms = {
    1: 'baumschule',
    2: 'galabau',
    3: 'garten',
    4: 'gartencenter',
    5: 'baumschulen',
    6: 'gartenbau',
    7: 'landschaftsbau',
    8: 'gmbh',
    9: 'ag'
}

# read data of currently palnned transport
def	read_transport_plan(xl_file):
    names = []
    file = pd.read_excel(xl_file)
    names = file['Cliente'].tolist()
    return (names)

# remove common terms from customer's name and measure Levenshtein ratio
def	fuzzy_match(str1, str2):
    str1 = str1.lower()
    str2 = str2.lower()
    for key in common_terms:
        str1.replace(common_terms[key], '')
        str2.replace(common_terms[key], '')
    return (Lev.ratio(str1, str2))

# retrieve adress in CSV from name
def	get_adress(name, customer_info):
    name = name.lower()
    max_similarity = 0
    customer_adress = {}
    for i, row in customer_info.iterrows():
        current_ratio = fuzzy_match(name, row['Acronimo'].strip())
        if current_ratio > 0.8 and current_ratio > max_similarity:
            max_similarity = current_ratio
            customer_adress['Address'] = row['Indirizzo'].strip()
            customer_adress['CAP'] = row['CAP'].strip()
            customer_adress['City']= row['Localita\''].strip()
            customer_adress['Country'] = row['Nazione'].strip()
            customer_adress['Ragione Sociale'] = row['Ragione Sociale 1'].strip() + row['Ragione Sociale 2'].strip()
            print(f'Trovato indirizzo di: {name}\n')
    if max_similarity == 0:
        print(f"Indirizzo di {name} non trovato! Inserire manualmente\n")
        return None
    return (customer_adress)

# we need coordinates for the trip call to OSM server
def get_coordinates(zip_code, country):
    coordinates = {}
    geodb = pg.Nominatim(country)
    coordinates['latitude'] = geodb.query_postal_code(zip_code)['latitude']
    coordinates['longitude'] = geodb.query_postal_code(zip_code)['longitude']
    return (coordinates)

# unnecessary
def get_zip_code(coord):
    geolocator = Nominatim(user_agent="matteo.pascale@gorinipiante.it", timeout=10)
    location = geolocator.reverse((coord['latitude'], coord[0]), exactly_one=True)
    if location and 'postcode' in location.raw['address']:
        return location.raw['address']['postcode']
    return None

# build a dictionary of the customers with: name, address, zip, coordinates
def	customer_dict():
    customers = {}
    customer_info = pd.read_csv('CLIENTI_FAT.csv', sep=';', on_bad_lines= 'skip')
    print('Analizzo file excel\n')
    for name in read_transport_plan('Costi.xlsx'):
        customers[name] = get_adress(name, customer_info)
        if customers[name]:
            customers[name]['Coordinates'] = get_coordinates(customers[name]['CAP'], customers[name]['Country'])
        time.sleep(2)
    return (customers)

# api call. Note source=first & roundtrip=false & geometries=polyline
def	osm_request(coord_list):
    coord_str = ";".join([f"{lon},{lat}" for lon, lat in coord_list])
    url = (
        f"http://router.project-osrm.org/trip/v1/driving/{coord_str}"
        "?source=first&roundtrip=false&geometries=polyline"
    )
    r = requests.get(url)
    data = r.json()
    return (data)

# get the polyline and the unloading order. The latter is constructed with waypoint indices
def shortest_path(adress_dict):
    coord_list = []
    coord_list.append((10.972482955970866, 43.91835625149449))
    for customer in adress_dict:
        if adress_dict[customer]:
            coord = (adress_dict[customer]['Coordinates']['longitude'], adress_dict[customer]['Coordinates']['latitude'])
            coord_list.append(coord)
    osm_response = osm_request(coord_list)
    unloading_order = []
    for wp in osm_response["waypoints"]:
        unloading_order.append(wp["waypoint_index"])
    return (unloading_order, osm_response)

# standard formulation for the request to Marta Transport    
def print_email(unloading_order, customer_dict):
    text = """Hello Marta,
 
This is the order for <DATE>
Price <PRICE>\n"""
    xlsx_file = pd.read_excel('Costi.xlsx')
    for i in range(1, len(unloading_order)):
        for stop in unloading_order:
            if stop == i:
                j = unloading_order.index(stop) - 1
                name = xlsx_file.iloc[j]['Cliente']
                text += f"{i}. {customer_dict[name]['Ragione Sociale']}, {customer_dict[name]['Address']}, {customer_dict[name]['CAP']}\n"
    with open("mail_marta.txt", "w") as f:
        f.write(text)
    print('Testo mail stampato in mail_marta.txt\n')

# decode the polyline and save an html map with folium
def print_route(osm_result):
    m = folium.Map()   
    folium.PolyLine(locations=polyline.decode(osm_result["trips"][0]["geometry"])).add_to(m)
    m.save('percorso.html')
    print('Percorso salvato nel file percorso.html\n')

def	main():
    customers = customer_dict()
    unloading_order, osm_result = shortest_path(customers)
    print_email(unloading_order, customers)
    print_route(osm_result)

if __name__ == "__main__":
    main()