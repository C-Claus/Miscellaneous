import os 
import lxml 
import uuid
import ifcopenshell

from lxml import etree as ET
from datetime import datetime
from collections import defaultdict
#from builtins import list


ifcfile = ifcopenshell.open('C:\\Users\\CClaus\\Desktop\\Flat 11\\ruimtemodel_flat_11.ifc')

products = ifcfile.by_type('IfcProduct')
zones = ifcfile.by_type('IfcZone')
spaces = ifcfile.by_type('IfcSpace')
building_stories = ifcfile.by_type('IfcBuildingStorey')

#create building storey list
building_storey_list = []

for i in building_stories:
    building_storey_list.append(i.Name)

now = datetime.now()
year = now.strftime("%Y")
month = now.strftime("%m")
day = now.strftime("%d")


def create_xml(file_name, building_storey, zone):
  
    root = ET.Element('SMARTVIEWSETS')
    doc = ET.SubElement(root, "SMARTVIEWSET")
 
    title = ET.SubElement(doc, "TITLE").text = "location_" + building_storey
    description = ET.SubElement(doc, "DESCRIPTION").text = "Automate Smart Views" 
    guid = ET.SubElement(doc, "GUID" ).text = str(uuid.uuid4())
    smartviews = ET.SubElement(doc, "SMARTVIEWS")
    

    for i in zone:
        smartview = ET.SubElement(smartviews, "SMARTVIEW")
        smartview_title = ET.SubElement(smartview, "TITLE").text = 'Building number: ' + str(i) #+ str(i.Name) 
        smartview_description = ET.SubElement(smartview, "DESCRIPTION")
        smartview_creator = ET.SubElement(smartview, "CREATOR").text = "C. Claus"
        smartview_creation_date = ET.SubElement(smartview, "CREATIONDATE").text = day + " " + month + " " + year
        smartview_modifier = ET.SubElement(smartview, "MODIFIER").text = "C. Claus"
        smartview_modification_date = ET.SubElement(smartview, "MODIFICATIONDATE").text = day + " " + month + " " + year
        smartview_guid = ET.SubElement(smartview, "GUID").text = str(uuid.uuid4())
        rules = ET.SubElement(smartview, "RULES")
 
        first_rule(smartview, rules,  building_storey)
        
        third_rule(smartview, rules, building_storey='building storey')
        
        second_rule(smartview, rules, zones=i)
        
     
      
        
    tree = ET.ElementTree(root)
    tree.write('bcsv_files/' + str(file_name), encoding="utf-8", xml_declaration=True, pretty_print=True)
   
    create_xml_declaration(file_path_xml='bcsv_files/' + str(file_name))
   
   
def first_rule(smartview, rules, building_storey): 
    
        rule = ET.SubElement(rules, "RULE")

        ifctype = ET.SubElement(rule, "IFCTYPE").text = 'Building Story'
        
        property = ET.SubElement(rule, "PROPERTY")
        property_name = ET.SubElement(property, "NAME").text= "Name"
        propertyset_name = ET.SubElement(property, "PROPERTYSETNAME").text = "Summary"
        property_type = ET.SubElement(property, "TYPE").text = "Summary"
        property_value_type = ET.SubElement(property, "VALUETYPE").text = "StringValue"
        property_unit = ET.SubElement(property, "UNIT").text = "None"
        
        
        condition = ET.SubElement(rule, "CONDITION")
        condition_type = ET.SubElement(condition, "TYPE").text = "StartsWith"
        condition_value = ET.SubElement(condition, "VALUE").text =  building_storey
        
        action = ET.SubElement(rule, "ACTION")
        action_type = ET.SubElement(action, "TYPE").text = "AddSetColored"
        
        r_color = ET.SubElement(action, "R").text = "145"
        g_color = ET.SubElement(action, "G").text = "145"
        b_color = ET.SubElement(action, "B").text = "145"
        
        
def second_rule(smartview, rules, zones):     
    
        #rules = ET.SubElement(smartview, "RULES")
        rule = ET.SubElement(rules, "RULE")

        ifctype = ET.SubElement(rule, "IFCTYPE").text = 'Zone'
        
        property = ET.SubElement(rule, "PROPERTY")
        property_name = ET.SubElement(property, "NAME").text= "Name"
        propertyset_name = ET.SubElement(property, "PROPERTYSETNAME").text = "Summary"
        property_type = ET.SubElement(property, "TYPE").text = "Summary"
        property_value_type = ET.SubElement(property, "VALUETYPE").text = "StringValue"
        property_unit = ET.SubElement(property, "UNIT").text = "None"
        
        
        condition = ET.SubElement(rule, "CONDITION")
        condition_type = ET.SubElement(condition, "TYPE").text = "StartsWith"
        condition_value = ET.SubElement(condition, "VALUE").text =  zones
        
        action = ET.SubElement(rule, "ACTION")
        action_type = ET.SubElement(action, "TYPE").text = "AddSetColored"
        
        r_color = ET.SubElement(action, "R").text = "255"
        g_color = ET.SubElement(action, "G").text = "0"
        b_color = ET.SubElement(action, "B").text = "0"   
        
def third_rule(smartview, rules, building_storey): 
    

        
        for b in building_storey_list:
            rule = ET.SubElement(rules, "RULE")
    
            ifctype = ET.SubElement(rule, "IFCTYPE").text = 'Building Story'
            
            property = ET.SubElement(rule, "PROPERTY")
            property_name = ET.SubElement(property, "NAME").text= "Name"
            propertyset_name = ET.SubElement(property, "PROPERTYSETNAME").text = "Summary"
            property_type = ET.SubElement(property, "TYPE").text = "Summary"
            property_value_type = ET.SubElement(property, "VALUETYPE").text = "StringValue"
            property_unit = ET.SubElement(property, "UNIT").text = "None"
            
            
            condition = ET.SubElement(rule, "CONDITION")
            condition_type = ET.SubElement(condition, "TYPE").text = "StartsWith"
            condition_value = ET.SubElement(condition, "VALUE").text =  b #uilding_storey
            
            action = ET.SubElement(rule, "ACTION")
            action_type = ET.SubElement(action, "TYPE").text = "AddSetTransparent"
        
            r_color = ET.SubElement(action, "R").text = "255"
            g_color = ET.SubElement(action, "G").text = "255"
            b_color = ET.SubElement(action, "B").text = "255"           
 
 

       
def create_xml_declaration(file_path_xml):
    
    xml_declaration =   """<?xml version="1.0"?>
    <bimcollabsmartviewfile>
    <version>5</version>
    <applicationversion>Win - Version: 3.1 (build 3.1.13.217)</applicationversion>
</bimcollabsmartviewfile>
    """
    
    with open(file_path_xml, 'r') as xml_file:
        xml_data = xml_file.readlines()
        
        
    xml_data[0] = xml_declaration + '\n'
    
    with open(file_path_xml, 'w') as xml_file:
        xml_file.writelines(xml_data)    
        
        
 
def get_building_storey_and_spaces_zone():    
    
    zone_list = []
    
    for building_storey in building_stories:
        for i in (building_storey.IsDecomposedBy):
            for j in (i.RelatedObjects):
                for k in (j.HasAssignments):
                    if k.is_a('IfcRelAssignsToGroup'):
                        zone_list.append([building_storey.Name, j.Name, k.RelatingGroup.Name])
                
    return zone_list  


def organize_zones_per_building_storey():

    zone_list = []
    
    #create building storey and zone list
    for i in get_building_storey_and_spaces_zone():
        for j in building_storey_list:
            if i[0].startswith(j): 
                (zone_list.append([j, i[2]]))
                
       
    #uniqify the list 
    unique_data = [list(x) for x in set(tuple(x) for x in zone_list)]   
    unique_zone_list = []
    
    for i in sorted(unique_data):
        unique_zone_list.append((i))
        
        
    #sort the zone list by building store
    zone_dict = defaultdict(list)
    
    for k, v in unique_zone_list:
        zone_dict[k].append(v)
        
    return zone_dict
    
    
    
for k, v in organize_zones_per_building_storey().items():
    create_xml(file_name='zone_filter_' + k + '.bcsv', building_storey=k, zone=v)  
    
    
    
        




