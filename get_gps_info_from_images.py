import os
from GPSPhoto import gpsphoto
from PIL import Image
import time 

folder = "photos_2020-02-05\\"

def get_date_taken(path_to_image):
    
    for i in os.listdir(path_to_image):
        print( Image.open(path_to_image + i)._getexif()[36867])


def get_gps_info(path_to_image):

    for i in os.listdir(path_to_image):
        data = gpsphoto.getGPSData(path_to_image + i)
        print(i, data['Latitude'], data['Longitude'])
        
        
def get_date_and_location(path_to_image):
    
    name_date_location_list = []
    
    for i in os.listdir(path_to_image):

        time_stamp = (Image.open(path_to_image + i)._getexif()[36867])

        data = gpsphoto.getGPSData(path_to_image + i)

        name_date_location_list.append([i,  time_stamp, data['Latitude'], data['Longitude']])

        
    return name_date_location_list


def sort_list_by_time():
    
    print (get_date_and_location(path_to_image=folder))



sort_list_by_time()

   