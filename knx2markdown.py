import zipfile
import xml.etree.ElementTree as ET
import os
import csv
import sys

import argparse

# Configuration


# Namespaces
NS = {'knx': 'http://knx.org/xml/project/23'}

# Strings / Translations
STRINGS = {
    'de': {
        'code': 'de-DE',
        'report_title': 'KNX Projekt Bericht',
        'stats_header': 'Projekt Statistiken',
        'stats_ga': 'Gruppenadressen',
        'stats_dev': 'Geräte',
        'bldg_structure': 'Gebäudestruktur',
        'unknown_device': 'Unbekanntes Gerät',
        'gas_header': 'Gruppenadressen',
        'table_addr': 'Adresse',
        'table_name': 'Name',
        'table_dpt': 'DPT',
        'table_main': 'Hauptgruppe',
        'table_mid': 'Mittelgruppe',
        'dev_header': 'Geräte',
        'dev_legend': '**Legende**:\n> * **Flags**: `A`=Adresse, `P`=Programm, `Pa`=Parameter, `G`=Gruppen, `K`=Konfig. [X] = Programmiert, [.] = Nicht programmiert, [?] = Unbekannt',
        'table_prodref': 'Produkt Ref',
        'table_status': 'Prog. Status',
        'conn_header': 'Detaillierte Geräteverbindungen',
        'conn_legend': '**Legende**:\n> * **Flags**: `K`=Kommunikation, `L`=Lesen, `S`=Schreiben, `Ü`=Übertragen, `A`=Aktualisieren\n> * **Verknüpfungen**: `[S]` = Sendende Adresse (Erste Verknüpfung wenn Ü-Flag aktiv)',
        'conn_obj': 'Obj',
        'conn_func': 'Funktion / Name',
        'conn_flags': 'Flags',
        'conn_links': 'Verknüpfte Gruppen',
        'param_header': 'Geräteparameter',
        'param_note': '**Hinweis**: Zeigt explizit konfigurierte Parameter aus der Projektdatei.',
        'param_none': '_Keine Parameterdaten gefunden (oder nicht geparst)._',
        'table_val': 'Wert',
        'table_raw': 'Raw',
        'desc': 'Beschreibung',
        # Translations for Types and Flags
        'types': {
            'Building': 'Gebäude',
            'Floor': 'Etage',
            'Room': 'Raum',
            'DistributionBoard': 'Verteiler',
            'Stairway': 'Treppenhaus',
            'Corridor': 'Flur',
            'Cabinet': 'Schrank',
        },
        'status_flag_map': {
            'Adr': 'A', 'Prg': 'P', 'Par': 'Pa', 'Grp': 'G', 'Cfg': 'K'
        },
        'conn_flag_map': {
            'C': 'K', 'R': 'L', 'W': 'S', 'T': 'Ü', 'U': 'A'
        }
    },
    'en': {
        'code': 'en-US',
        'report_title': 'KNX2Markdown Report',
        'stats_header': 'Project Statistics',
        'stats_ga': 'Group Addresses',
        'stats_dev': 'Devices',
        'bldg_structure': 'Building Structure',
        'unknown_device': 'Unknown Device',
        'gas_header': 'Group Addresses',
        'table_addr': 'Address',
        'table_name': 'Name',
        'table_dpt': 'DPT',
        'table_main': 'Main Group',
        'table_mid': 'Middle Group',
        'dev_header': 'Devices',
        'dev_legend': '**Prog. Status**: `Ad`=Address, `Pr`=Program, `Pa`=Parameters, `Gr`=Groups, `Cf`=Config. [X] = Programmed, [.] = Not programmed, [?] = Unknown',
        'table_prodref': 'Product Ref',
        'table_status': 'Prog. Status',
        'conn_header': 'Detailed Device Connections',
        'conn_legend': '**Legend**:\n> * **Flags**: `C`=Comm, `R`=Read, `W`=Write, `T`=Transmit, `U`=Update\n> * **Links**: `[S]` = Sending Address (First link if T-Flag is active)',
        'conn_obj': 'Obj',
        'conn_func': 'Function / Name',
        'conn_flags': 'Flags',
        'conn_links': 'Linked Groups',
        'param_header': 'Device Parameters',
        'param_note': '**Note**: Shows explicitly configured parameters found in the project file.',
        'param_none': '_No parameter data found (or not parsed)._',
        'table_val': 'Value',
        'table_raw': 'Raw',
        'desc': 'Description',
        # Translations for Types and Flags
        'types': {
            'Building': 'Building',
            'Floor': 'Floor',
            'Room': 'Room',
            'DistributionBoard': 'DistributionBoard',
            'Stairway': 'Stairway',
            'Corridor': 'Corridor',
            'Cabinet': 'Cabinet',
        },
        'status_flag_map': {
            'Adr': 'Ad', 'Prg': 'Pr', 'Par': 'Pa', 'Grp': 'Gr', 'Cfg': 'Cf'
        },
        'conn_flag_map': {
            'C': 'C', 'R': 'R', 'W': 'W', 'T': 'T', 'U': 'U'
        }
    }
}


def extract_project_xml(knxproj_path):
    """Extracts 0.xml (actual project data) from the .knxproj zip archive."""
    with zipfile.ZipFile(knxproj_path, 'r') as z:
        # iterate to find the correct path, usually P-xxxx/0.xml
        for filename in z.namelist():
            if filename.endswith('0.xml') and len(filename.split('/')) == 2:
                print(f"Extracting {filename}...")
                return z.read(filename)
    raise FileNotFoundError("Could not find project data (0.xml) in .knxproj")

def parse_group_addresses(root):
    """Parses Group Addresses from the XML tree."""
    # Structure: GroupAddresses -> GroupRanges -> GroupRange -> GroupRange -> GroupAddress
    addresses = []
    
    # helper to safely get attribute
    def get_attr(el, name):
        return el.attrib.get(name, "")

    # Iterate through Installations -> Installation
    for installation in root.findall('.//knx:Installation', NS):
        gas_node = installation.find('knx:GroupAddresses', NS)
        if gas_node is None:
            continue
            
        for range1 in gas_node.findall('knx:GroupRanges/knx:GroupRange', NS):
            range1_name = get_attr(range1, 'Name')
            
            for range2 in range1.findall('knx:GroupRange', NS):
                range2_name = get_attr(range2, 'Name')
                
                for ga in range2.findall('knx:GroupAddress', NS):
                    addr_val = int(get_attr(ga, 'Address'))
                    # Convert integer address to x/y/z format (3-level)
                    main = (addr_val >> 11) & 0x1F
                    middle = (addr_val >> 8) & 0x07
                    sub = addr_val & 0xFF
                    addr_str = f"{main}/{middle}/{sub}"
                    
                    addresses.append({
                        'Id': ga.attrib.get('Id'),
                        'Address': addr_str,
                        'Name': get_attr(ga, 'Name'),
                        'Description': get_attr(ga, 'Description'),
                        'DPT': get_attr(ga, 'DatapointType'),
                        'MainGroup': range1_name,
                        'MiddleGroup': range2_name
                    })
    return addresses

def get_ga_lookup(addresses):
    """Creates a lookup dictionary {Id: {'Address': 'x/y/z', 'Name': '...'}, ...}"""
    lookup = {}
    for ga in addresses:
        if 'Id' in ga:
            lookup[ga['Id']] = ga
            
            # Also map the short ID 'GA-xxx' used in Links
            if '_GA-' in ga['Id']:
                short_id = 'GA-' + ga['Id'].split('_GA-')[-1]
                lookup[short_id] = ga
    return lookup

def parse_devices(root, product_lookup, ga_lookup, comobj_lookup, product_app_map, param_lookup, param_types):
    """Parses Topology/Devices and their Parameters."""
    print(f"DEBUG: Parsing devices for parameters...")
    devices = []
    device_lookup = {}
    
    # Iterate through Topology recursively to find all DeviceInstance nodes
    for installation in root.findall('.//knx:Installation', NS):
        topo_node = installation.find('knx:Topology', NS)
        if topo_node is None:
            continue
            
        for area in topo_node.findall('knx:Area', NS):
            area_addr = area.attrib.get('Address', '0')
            
            for line in area.findall('knx:Line', NS):
                line_addr = line.attrib.get('Address', '0')
                
                for dev in line.findall('.//knx:DeviceInstance', NS):
                    dev_addr = dev.attrib.get('Address', '0')
                    full_addr = f"{area_addr}.{line_addr}.{dev_addr}"
                    dev_id = dev.attrib.get('Id', '')
                    product_ref = dev.attrib.get('ProductRefId', '')
                    app_prog_id = product_app_map.get(product_ref, '')
                    
                    # Programming Status Flags
                    status = {
                        'Adr': dev.attrib.get('IndividualAddressLoaded', 'false') == 'true',
                        'Prg': dev.attrib.get('ApplicationProgramLoaded', 'false') == 'true',
                        'Par': dev.attrib.get('ParametersLoaded', 'false') == 'true',
                        'Grp': dev.attrib.get('CommunicationPartLoaded', 'false') == 'true',
                        'Cfg': dev.attrib.get('MediumConfigLoaded', 'false') == 'true'
                    }
                    com_objects = []
                    com_refs = dev.find('knx:ComObjectInstanceRefs', NS)
                    if com_refs is not None:
                        for ref in com_refs:
                            links_str = ref.attrib.get('Links', '')
                            ref_id = ref.attrib.get('RefId') # e.g., O-0_R-0
                            
                            # Resolve Name
                            obj_name = ""
                            base_id = ref_id.split('_')[0] if ref_id else ""
                            
                            if base_id and app_prog_id:
                                full_key = f"{app_prog_id}_{base_id}"
                                if full_key in comobj_lookup:
                                    obj_name = comobj_lookup[full_key]
                            
                            # Fallback if full key failed
                            if not obj_name and ref_id in comobj_lookup:
                                obj_name = comobj_lookup[ref_id]
                            
                            if not obj_name and base_id in comobj_lookup:
                                obj_name = comobj_lookup[base_id]

                            linked_gas = []
                            if links_str:
                                for ga_id in links_str.split(): 
                                    if ga_id in ga_lookup:
                                        linked_gas.append(ga_lookup[ga_id]['Address'])
                                    else:
                                        linked_gas.append(ga_id) # Keep raw ID if not found

                            if linked_gas or obj_name or ref_id: # Include even if just info
                                flags = []
                                if ref.attrib.get('ReadFlag') == 'Enabled': flags.append('R')
                                if ref.attrib.get('WriteFlag') == 'Enabled': flags.append('W')
                                if ref.attrib.get('TransmitFlag') == 'Enabled': flags.append('T')
                                if ref.attrib.get('UpdateFlag') == 'Enabled': flags.append('U')
                                
                                com_objects.append({
                                    'RefId': ref_id,
                                    'Name': obj_name,
                                    'Links': linked_gas,
                                    'Flags': "".join(flags)
                                })

                    # --- 2. Extract Parameters ---
                    parameters = []
                    param_refs = dev.find('knx:ParameterInstanceRefs', NS)
                    if param_refs is not None:
                        # Iterate ParameterInstanceRef
                        for p_ref in param_refs: 
                            val = p_ref.attrib.get('Value')
                            p_ref_id = p_ref.attrib.get('RefId') 
                            
                            if not p_ref_id: continue

                            # Resolve Name/Text from param_lookup
                            # ID usually: M-0083_A-...._P-123_R-456
                            # Base Parameter ID: M-0083_A-...._P-123
                            
                            # Try to strip _R- suffix
                            base_id = ""
                            if '_R-' in p_ref_id:
                                base_id = p_ref_id.rsplit('_R-', 1)[0]
                            else:
                                base_id = p_ref_id
                            
                            p_def = param_lookup.get(base_id)
                            if p_def:
                                p_text = p_def['Text']
                                p_name = p_def.get('Name', '')
                                
                                # Fix cryptic PID names (Couplers)
                                if p_text.startswith('PID_') and p_name:
                                    # Name is like "PID_MAIN_LCCONFIG_PHYFILTER (52)"
                                    # Clean up: strip PID_ prefix and (52) suffix
                                    clean_name = p_name
                                    if clean_name.startswith('PID_'):
                                        # Remove common prefixes
                                        for prefix in ['PID_MAIN_LCCONFIG_', 'PID_SUB_LCCONFIG_', 'PID_MAIN_', 'PID_SUB_', 'PID_']:
                                            if clean_name.startswith(prefix):
                                                clean_name = clean_name.replace(prefix, '')
                                                break
                                    
                                    # Remove ID in brackets if present e.g. " (52)"
                                    if '(' in clean_name:
                                        clean_name = clean_name.split('(')[0].strip()
                                        
                                    p_text = clean_name
                                    
                                p_type = p_def['Type']
                                p_suffix = p_def['Suffix']
                                
                                # Resolve Value (Enum or Unit)
                                display_value = val
                                if p_type in param_types:
                                    if val in param_types[p_type]:
                                        display_value = param_types[p_type][val]
                                
                                if p_suffix and display_value == val:
                                     display_value = f"{val} {p_suffix}"
                                
                                parameters.append({
                                    'Name': p_text,
                                    'Value': display_value,
                                    'RawValue': val
                                })

                    dev_obj = {
                        'Id': dev_id,
                        'Address': full_addr,
                        'Name': dev.attrib.get('Name', ''),
                        'ProductRef': product_ref,
                        'Description': dev.attrib.get('Description', ''),
                        'ComObjects': com_objects,
                        'Parameters': parameters,
                        'Status': status
                    }
                    devices.append(dev_obj)
                    if dev_id:
                        device_lookup[dev_id] = dev_obj

    return devices, device_lookup

def parse_hardware_catalog(knxproj_path, language_code='de-DE'):
    """Parses M-*/Hardware.xml files to build lookup tables."""
    product_lookup = {}
    product_app_map = {}
    
    with zipfile.ZipFile(knxproj_path, 'r') as z:
        for filename in z.namelist():
            if filename.endswith('Hardware.xml') and filename.count('/') == 1:
                # e.g., M-0083/Hardware.xml
                try:
                    xml_content = z.read(filename)
                    root = ET.fromstring(xml_content)
                    
                    ns_match = root.tag.split('}')
                    ns = {}
                    if len(ns_match) > 1:
                         ns = {'knx': ns_match[0].strip('{')}
                    
                    # Iterate Hardware elements
                    for hw in root.findall('.//knx:Hardware', ns) if ns else root.findall('.//Hardware'):
                        # Find App Program Ref
                        app_ref = ""
                        h2p = hw.find('.//knx:Hardware2Program', ns) if ns else hw.find('.//Hardware2Program')
                        if h2p is not None:
                            apr = h2p.find('knx:ApplicationProgramRef', ns) if ns else h2p.find('ApplicationProgramRef')
                            if apr is not None:
                                app_ref = apr.attrib.get('RefId')
                        
                        # Find Products and map them
                        products = hw.findall('.//knx:Product', ns) if ns else hw.findall('.//Product')
                        for product in products:
                            pid = product.attrib.get('Id')
                            text = product.attrib.get('Text')
                            if pid and text:
                                product_lookup[pid] = text
                            if pid and app_ref:
                                product_app_map[pid] = app_ref
                              
                    # Search for Translations (German preferred)
                    languages = root.findall('.//knx:Language', ns) if ns else root.findall('.//Language')
                    for lang in languages:
                        if lang.attrib.get('Identifier') == language_code:
                            for unit in lang.findall('knx:TranslationUnit', ns) if ns else lang.findall('TranslationUnit'):
                                ref_id = unit.attrib.get('RefId')
                                if ref_id in product_lookup:
                                    # Look for Text content
                                    for element in unit.findall('knx:TranslationElement', ns) if ns else unit.findall('TranslationElement'):
                                         for trans in element.findall('knx:Translation', ns) if ns else element.findall('Translation'):
                                             if trans.attrib.get('AttributeName') == 'Text':
                                                 product_lookup[ref_id] = trans.attrib.get('Text')

                except Exception as e:
                    print(f"Error parsing {filename}: {e}")
                    
    return product_lookup, product_app_map

def parse_application_programs(knxproj_path, language_code='de-DE'):
    """
    Parses M-*/M-*.xml files to build:
    1. ComObject RefId -> {Text, FunctionText} lookup
    2. Parameter lookup: {AppId_PId: {Text, Type, Suffix}}
    3. ParameterType (Enum) lookup: {TypeId: {Value: Text}}
    """
    comobj_lookup = {}
    param_lookup = {}
    param_types = {}
    
    with zipfile.ZipFile(knxproj_path, 'r') as z:
        for filename in z.namelist():
            # Match Application Program files: M-0083/M-0083_A-....xml
            if filename.count('/') == 1 and filename.endswith('.xml') and '_A-' in filename:
                try:
                    xml_content = z.read(filename)
                    root = ET.fromstring(xml_content)
                    
                    ns_match = root.tag.split('}')
                    ns = {}
                    if len(ns_match) > 1:
                         ns = {'knx': ns_match[0].strip('{')}
                    
                    # 1. Parse Parameter Types (Enumerations)
                    for ptype in root.findall('.//knx:ParameterType', ns) if ns else root.findall('.//ParameterType'):
                        ptid = ptype.attrib.get('Id')
                        enums = {}
                        
                        type_res = ptype.find('.//knx:TypeRestriction', ns) if ns else ptype.find('.//TypeRestriction')
                        if type_res is not None:
                            for enum in type_res.findall('.//knx:Enumeration', ns) if ns else type_res.findall('.//Enumeration'):
                                val = enum.attrib.get('Value')
                                text = enum.attrib.get('Text')
                                if val and text:
                                    enums[val] = text
                        
                        if enums:
                            param_types[ptid] = enums

                    # 2. Parse Parameters
                    for param in root.findall('.//knx:Parameter', ns) if ns else root.findall('.//Parameter'):
                        pid = param.attrib.get('Id')
                        ptext = param.attrib.get('Text', '')
                        pname = param.attrib.get('Name', '') # Internal name often has more detail
                        ptype = param.attrib.get('ParameterType')
                        psuffix = param.attrib.get('Suffix', '') 
                        
                        param_lookup[pid] = {
                            'Text': ptext,
                            'Name': pname,
                            'Type': ptype,
                            'Suffix': psuffix
                        }

                    # 3. Parse ComObjects
                    for co in root.findall('.//knx:ComObject', ns) if ns else root.findall('.//ComObject'):
                        ref_id = co.attrib.get('Id')
                        text = co.attrib.get('Text')
                        function_text = co.attrib.get('FunctionText')
                        
                        if ref_id:
                            name = text if text else (function_text if function_text else "")
                            comobj_lookup[ref_id] = name
                            
                    # 4. Translations
                    languages = root.findall('.//knx:Language', ns) if ns else root.findall('.//Language')
                    for lang in languages:
                        if lang.attrib.get('Identifier') == language_code:
                            for unit in lang.findall('knx:TranslationUnit', ns) if ns else lang.findall('TranslationUnit'):
                                ref_id = unit.attrib.get('RefId')
                                
                                # ComObject Translation
                                if ref_id in comobj_lookup:
                                    for element in unit.findall('knx:TranslationElement', ns) if ns else unit.findall('TranslationElement'):
                                         for trans in element.findall('knx:Translation', ns) if ns else element.findall('Translation'):
                                             if trans.attrib.get('AttributeName') == 'Text':
                                                 comobj_lookup[ref_id] = trans.attrib.get('Text')
                                             elif trans.attrib.get('AttributeName') == 'FunctionText' and not comobj_lookup[ref_id]:
                                                 comobj_lookup[ref_id] = trans.attrib.get('Text')

                                # Parameter Translation
                                if ref_id in param_lookup:
                                     for element in unit.findall('knx:TranslationElement', ns) if ns else unit.findall('TranslationElement'):
                                         for trans in element.findall('knx:Translation', ns) if ns else element.findall('Translation'):
                                             if trans.attrib.get('AttributeName') == 'Text':
                                                 param_lookup[ref_id]['Text'] = trans.attrib.get('Text')

                except Exception as e:
                    print(f"Error parsing AppProgram {filename}: {e}")
                    
    return comobj_lookup, param_lookup, param_types

def parse_locations(root, device_lookup):
    """Parses Building/Space structure from the XML tree."""
    structure = []

    def parse_space_recursive(element, level=0):
        spaces = []
        for space in element.findall('knx:Space', NS):
            name = space.attrib.get('Name', '')
            stype = space.attrib.get('Type', '')
            
            # Find functions
            funcs = []
            for func in space.findall('knx:Function', NS):
                funcs.append(func.attrib.get('Name', ''))
            
            # Find Device References
            linked_devices = []
            for dev_ref in space.findall('knx:DeviceInstanceRef', NS):
                ref_id = dev_ref.attrib.get('RefId')
                if ref_id in device_lookup:
                    linked_devices.append(device_lookup[ref_id])
            
            # Recurse
            children = parse_space_recursive(space, level + 1)
            
            spaces.append({
                'Name': name,
                'Type': stype,
                'Functions': funcs,
                'Devices': linked_devices,
                'Children': children
            })
        return spaces

    for installation in root.findall('.//knx:Installation', NS):
        loc_node = installation.find('knx:Locations', NS)
        if loc_node is not None:
            structure.extend(parse_space_recursive(loc_node))
            
    return structure



def generate_markdown(gas, devices, locations, filename, lang='de'):
    """
    Generates the Markdown report in 'Expert' mode.
    Combines readability (names) with detailed technical info (flags, direction).
    """
    S = STRINGS[lang]
    # Create Address -> Name lookup
    addr_to_name = {g['Address']: g['Name'] for g in gas}

    with open(filename, 'w', encoding='utf-8') as f:
        f.write(f"# {S['report_title']}\n\n")
        
        f.write(f"## {S['stats_header']}\n")
        f.write(f"- **{S['stats_ga']}**: {len(gas)}\n")
        f.write(f"- **{S['stats_dev']}**: {len(devices)}\n\n")
        
        f.write(f"## {S['bldg_structure']}\n")
        
        def write_structure(spaces, level=0):
            indent = "  " * level
            for space in spaces:
                icon = ""
                if space['Type'] == 'Building': icon = "🏠 "
                elif space['Type'] == 'Floor': icon = "🪜 "
                elif space['Type'] == 'Room': icon = "🚪 "
                
                # Translate Type
                type_str = space['Type']
                if type_str in S['types']:
                    type_str = S['types'][type_str]
                
                f.write(f"{indent}- {icon}**{space['Name']}** ({type_str})\n")
                
                if space['Functions']:
                    for func in space['Functions']:
                        f.write(f"{indent}  - ⚙️ {func}\n")

                if space['Devices']:
                    for dev in space['Devices']:
                        display_name = dev['Name']
                        if not display_name:
                             display_name = dev.get('CatalogName', '') 
                        if not display_name:
                             display_name = dev.get('Description', '') 
                        if not display_name:
                             display_name = S['unknown_device']
                             
                        f.write(f"{indent}  - 🔌 **{display_name}** ({dev['Address']}) - {dev['ProductRef']}\n")
                
                write_structure(space['Children'], level + 1)
        
        write_structure(locations)
        f.write("\n")

        f.write(f"## {S['gas_header']}\n")
        f.write(f"| {S['table_addr']} | {S['table_name']} | {S['table_dpt']} | {S['table_main']} | {S['table_mid']} |\n")
        f.write("|---|---|---|---|---|\n")
        for ga in sorted(gas, key=lambda x: [int(p) for p in x['Address'].split('/')]):
            dpt = ga['DPT'] if ga['DPT'] else "-"
            f.write(f"| {ga['Address']} | {ga['Name']} | {dpt} | {ga['MainGroup']} | {ga['MiddleGroup']} |\n")
            
        f.write(f"\n## {S['dev_header']}\n")
        f.write(f"> {S['dev_legend']}\n\n")
        f.write(f"| {S['table_addr']} | {S['table_name']} | {S['table_prodref']} | {S['table_status']} |\n")
        f.write("|---|---|---|---|\n")
        for dev in sorted(devices, key=lambda x: [int(p) for p in x['Address'].split('.')]):
             cat_name = dev.get('CatalogName', '-')
             
             # Format Status
             # Ad Pr Pa Gr Cf
             s = dev.get('Status', {})
             
             def flag(key, code_key):
                 return S['status_flag_map'][code_key] if s.get(key) else "."
                 
             st_str = f"`{flag('Adr','Adr')}{flag('Prg','Prg')}{flag('Par','Par')}{flag('Grp','Grp')}{flag('Cfg','Cfg')}`"
             
             f.write(f"| {dev['Address']} | {dev['Name']} | {cat_name} | {st_str} |\n")

        f.write(f"\n## {S['conn_header']}\n")
        f.write(f"> {S['conn_legend']}\n\n")
        
        for dev in sorted(devices, key=lambda x: [int(p) for p in x['Address'].split('.')]):
            if 'ComObjects' in dev and dev['ComObjects']:
                cat_name = dev.get('CatalogName', dev['ProductRef'])
                f.write(f"### 🔌 {dev['Address']} - {dev['Name']} ({cat_name})\n")
                if dev['Description']:
                    f.write(f"- **{S['desc']}**: {dev['Description']}\n")
                
                f.write(f"| {S['conn_obj']} | {S['conn_func']} | {S['conn_flags']} | {S['conn_links']} |\n")
                f.write("|---|---|---|---|\n")
                
                for co in dev['ComObjects']:
                    # Clean up RefId
                    obj_num = co['RefId']
                    if 'O-' in obj_num:
                        obj_num = obj_num.split('O-')[-1].split('_')[0]
                    
                    name = co['Name'] if co['Name'] else "-"
                    
                    raw_flags = co['Flags'] # e.g. "RWT"
                    
                    f_c = S['conn_flag_map']['C'] if co['Links'] else "."
                    f_r = S['conn_flag_map']['R'] if "R" in raw_flags else "."
                    f_w = S['conn_flag_map']['W'] if "W" in raw_flags else "."
                    f_t = S['conn_flag_map']['T'] if "T" in raw_flags else "."
                    f_u = S['conn_flag_map']['U'] if "U" in raw_flags else "."
                    
                    flags_visual = f"`{f_c}{f_r}{f_w}{f_t}{f_u}`"
                    
                    # Process Links with Direction
                    links_formatted = "-"
                    if co['Links']:
                        formatted_list = []
                        is_sending = "T" in raw_flags
                        
                        for i, link in enumerate(co['Links']):
                            link_name = addr_to_name.get(link, "")
                            
                            prefix = "" 
                            if i == 0 and is_sending:
                                prefix = "**[S]** "
                            
                            if link_name:
                                formatted_list.append(f"{prefix}`{link}` {link_name}")
                            else:
                                formatted_list.append(f"{prefix}`{link}`")
                                
                        links_formatted = "<br>".join(formatted_list)
                    
                    f.write(f"| {obj_num} | {name} | {flags_visual} | {links_formatted} |\n")
                f.write("\n")

        # --- Device Parameters Section ---
        f.write(f"\n## {S['param_header']}\n")
        f.write(f"> {S['param_note']}\n\n")
        
        # Only devices with parameters
        devices_with_params = [d for d in devices if 'Parameters' in d and d['Parameters']]
        
        if not devices_with_params:
            f.write(f"{S['param_none']}\n")
        
        for dev in sorted(devices_with_params, key=lambda x: [int(p) for p in x['Address'].split('.')]):
            cat_name = dev.get('CatalogName', dev['ProductRef'])
            f.write(f"### ⚙️ {dev['Address']} - {dev['Name']} ({cat_name})\n")
            f.write(f"| Parameter | {S['table_val']} | {S['table_raw']} |\n")
            f.write("|---|---|---|\n")
            
            for p in dev['Parameters']:
                f.write(f"| {p['Name']} | **{p['Value']}** | `{p['RawValue']}` |\n")
            f.write("\n")

    print(f"Markdown report generated: {filename}")

def main():
    parser = argparse.ArgumentParser(description="Convert KNX project to Markdown report.")
    parser.add_argument("file", nargs='?', help="Path to .knxproj file")
    parser.add_argument("--lang", choices=['de', 'en'], default='de', help="Output language (de/en), default: de")
    args = parser.parse_args()
    
    # 1. Determine Input File
    knxproj_path = ""
    if args.file:
        knxproj_path = args.file
    else:
        # Auto-detect .knxproj files in current directory
        # Exclude known output/temp files if any, though extension differs
        files = [f for f in os.listdir('.') if f.endswith('.knxproj')]
        if len(files) == 1:
            knxproj_path = files[0]
            print(f"Auto-detected project file: {knxproj_path}")
        elif len(files) > 1:
            print("Multiple .knxproj files found:")
            for i, f in enumerate(files):
                print(f"{i + 1}: {f}")
            print("Please specify one as an argument.")
            sys.exit(1)
        else:
            print("No .knxproj file found in current directory. Please provide path as argument.")
            sys.exit(1)

    if not os.path.exists(knxproj_path):
        print(f"Error: File '{knxproj_path}' not found.")
        sys.exit(1)

    # 2. Determine Output Filename
    # Example: "MyProject.knxproj" -> "MyProject_Export.md"
    base_name = os.path.splitext(os.path.basename(knxproj_path))[0]
    output_file = f"{base_name}.md"

    print(f"Processing '{knxproj_path}' (Language: {args.lang})...")
    
    # Selected language code for XML lookup
    lang_code = STRINGS[args.lang]['code']

    # 3. Read XML
    try:
        xml_content = extract_project_xml(knxproj_path)
        root = ET.fromstring(xml_content)
    except Exception as e:
        print(f"Error reading project XML: {e}")
        return
    
    # 4. Parse Hardware Catalog & Application Programs
    print("Parsing Hardware Catalog (this might take a moment)...")
    product_lookup, product_app_map = parse_hardware_catalog(knxproj_path, lang_code)
    
    print("Parsing Application Programs for ComObject detailed names...")
    comobj_lookup, param_lookup, param_types = parse_application_programs(knxproj_path, lang_code)
    
    # 5. Parse Data
    print("Parsing Group Addresses and Devices...")
    gas = parse_group_addresses(root)
    ga_lookup = get_ga_lookup(gas)
    
    devices, device_lookup = parse_devices(root, product_lookup, ga_lookup, comobj_lookup, product_app_map, param_lookup, param_types)
    
    # Enrich devices with Catalog Names
    for dev in devices:
        ref_id = dev['ProductRef']
        if ref_id in product_lookup:
            dev['CatalogName'] = product_lookup[ref_id]
            
    locations = parse_locations(root, device_lookup)
    
    # 6. Generate Output
    print(f"Generating report: {output_file}")
    generate_markdown(gas, devices, locations, output_file, args.lang)
    print("Done!")

if __name__ == "__main__":
    main()
