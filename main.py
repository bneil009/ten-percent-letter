import gspread
import creds
import requests
import json
import os
from docxtpl import DocxTemplate

#Login to google account services - pass api key from json file to connect python to google sheet ********
login = gspread.service_account(filename="service_account.json")
sheet_name = login.open("HOA")

tab_lookup = sheet_name.worksheet("10 Percent")
state = str(tab_lookup.acell("D4").value)
Hoa_name = str(tab_lookup.acell("B7").value)
Date = str(tab_lookup.acell("H7").value)
name = str(tab_lookup.acell("D7").value)
arrayCount = str(tab_lookup.acell("J2").value)
arrayCount = int(arrayCount)

if state == 'TX':
    state = f"""
Also, here is a short excerpt from the Texas Solar Rights that refers to this issue. “The law also 
stipulates that the HOA may designate where the solar device should be located on a roof,
unless a homeowner can show that the designation negatively impacts the performance
of the solar energy device and an alternative location would increase production by
more than 10%. To show this, the law requires that modeling tools provided by the
National Renewable Laboratory (NREL) be used.” 

While not specified by name in the law, one of NREL’s available tools that can accomplish this is called PVWatts Calculator.
http://programs.dsireusa.org/system/program/detail/4880"""
elif state == 'CO':
    state = f"""
Also, here is a short excerpt from the Colorado House Bill that refers to this issue.
"Section 2 of the act adds specificity to the requirements that HOAs allow installation
of renewable energy generation devices (e.g solar panels) subject to reasonable
aesthetic guidelines by requiring approval or denial of a completed application
within 60 days and requiring approval if imposition of the aesthetic
guidelines would result in more than a 10% reduction in efficiency or a 10% increase in price

https://leg.colorado.gov/sites/default/files/2021a_1229_signed.pdf"""
else:
    pass

def arrayOne():

    Quantity = str(tab_lookup.acell("M5").value)
    Quantity_2 = str(tab_lookup.acell("N5").value)
    Old_tilt = str(tab_lookup.acell("M2").value)
    Old_azimuth = str(tab_lookup.acell("M3").value)
    Old_direction = str(tab_lookup.acell("M6").value)
    New_direction = str(tab_lookup.acell("N6").value)
    New_tilt = str(tab_lookup.acell("N2").value)
    New_azimuth = str(tab_lookup.acell("N3").value)
    Mod_watt = str(tab_lookup.acell("C10").value)
    Address = str(tab_lookup.acell("B13:C13").value)
    Losses_o = str(tab_lookup.acell("M4").value)
    Losses_n = str(tab_lookup.acell("N4").value)
    Module_type = "1"
    Array_type = "1"
    System_capacity = 1
    System_capacity_2 = 1
    # ^^^ Google sheet values - checked and stringed ready to pass into docxtpl and calculations ^^^ *********

    # Calculating System_capacity ****************************************************************************
    Quantity = int(Quantity)
    Quantity_2 = int(Quantity_2)

    if Mod_watt == "SPR-M435-H-AC":
        System_capacity = Quantity * .435
        System_capacity_2 = Quantity_2 * .435
    elif Mod_watt == "SPR-M425-H-AC":
        System_capacity = Quantity * .425
        System_capacity_2 = Quantity_2 * .425
    elif Mod_watt == "SPR-A420-AC":
        System_capacity = Quantity * .420
        System_capacity_2 = Quantity_2 * .420
    elif Mod_watt == "SPR-A415-AC":
        System_capacity = Quantity * .415
        System_capacity_2 = Quantity_2 * .415
    elif Mod_watt == "SPR-A410-AC":
        System_capacity = Quantity * .410
        System_capacity_2 = Quantity_2 * .410
    elif Mod_watt == "JKM410M-72HL-V G2 410W":
        System_capacity = Quantity * .410
        System_capacity_2 = Quantity_2 * .410
    elif Mod_watt == "SPR-A400-BLK-AC":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-A400-BLK":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-A400-AC":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-U400-BLK":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-X22-370-AC":
        System_capacity = Quantity * .370
        System_capacity_2 = Quantity_2 * .370
    elif Mod_watt == "SPR-X22-360-AC":
        System_capacity = Quantity * .360
        System_capacity_2 = Quantity_2 * .360
    elif Mod_watt == "SPR-X22-360":
        System_capacity = Quantity * .360
        System_capacity_2 = Quantity_2 * .360
    elif Mod_watt == "SPR-X21-350-BLK-AC":
        System_capacity = Quantity * .350
        System_capacity_2 = Quantity_2 * .350
    elif Mod_watt == "SPR-E20-327-AC":
        System_capacity = Quantity * .327
        System_capacity_2 = Quantity_2 * .327
    elif Mod_watt == "SPR-E20-327":
        System_capacity = Quantity * .327
        System_capacity_2 = Quantity_2 * .327
    elif Mod_watt == "SPR-E19-320-AC":
        System_capacity = Quantity * .320
        System_capacity_2 = Quantity_2 * .320
    else:
        pass

    System_capacity = str(System_capacity)
    System_capacity_2 = str(System_capacity_2)

    # *********************************************************************************************************

    # setting string variables for NREL/PVWATTS parameters ****************************************************

    Address = Address.replace(" ", "%20").strip()
    Address = ("address=" + Address + "&")
    New_tilt = ("tilt=" + New_tilt + "&")
    New_azimuth = ("azimuth=" + New_azimuth + "&")
    Old_tilt = ("tilt=" + Old_tilt + "&")
    Old_azimuth = ("azimuth=" + Old_azimuth + "&")
    Losses_o = ("losses=" + Losses_o + "&")
    Losses_n = ("losses=" + Losses_n + "&")
    Module_type = ("module_type=" + Module_type + "&")
    Array_type = ("array_type=" + Array_type + "&")
    System_capacity = ("system_capacity=" + System_capacity + "&")
    System_capacity_2 = ("system_capacity=" + System_capacity_2 + "&")

    API_PARAM = "&api_key="
    OLD_QUERY = Address + Old_tilt + Old_azimuth + Losses_o + Module_type + Array_type + System_capacity + API_PARAM + creds.API_KEY
    NEW_QUERY = Address + New_tilt + New_azimuth + Losses_n + Module_type + Array_type + System_capacity_2 + API_PARAM + creds.API_KEY

    # parameters set for NREL/PVW API connection, preparing to make get call & parse data after 200 response *

    # preforming first requests to NREL/PVW API -- Looks a little jank but it works **************************
    response = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + API_PARAM + creds.API_KEY)
    BASE_URL = "https://developer.nrel.gov/api/pvwatts/v6.json?"

    JSON_LINK_ORIGINAL = (BASE_URL + OLD_QUERY)
    JSON_LINK_NEW = (BASE_URL + NEW_QUERY)
    data_original = requests.get(JSON_LINK_ORIGINAL)
    data_new = requests.get(JSON_LINK_NEW)

    print(JSON_LINK_ORIGINAL)
    print(JSON_LINK_NEW)
    print(data_original.status_code)
    print(data_new.status_code)

    # finished api request should have 200 response in console / printed *************************************

    # had a lot of trouble with the json formatting from NREL/PVW had to parse this way to get annual total **
    content = data_original.text
    data_original = json.loads(content)
    content = data_new.text
    data_new = json.loads(content)

    ac_monthly_original = data_original.get('outputs')
    ac_monthly_new = data_new.get('outputs')
    dict.items(ac_monthly_original)
    dict.items(ac_monthly_new)

    del [ac_monthly_original['ac_monthly']]
    del [ac_monthly_original['poa_monthly']]
    del [ac_monthly_original['solrad_monthly']]
    del [ac_monthly_original['dc_monthly']]
    del [ac_monthly_original['solrad_annual']]
    del [ac_monthly_original['capacity_factor']]
    ac_monthly_original = str(ac_monthly_original)
    ac_monthly_original = ac_monthly_original[14:]

    del [ac_monthly_new['ac_monthly']]
    del [ac_monthly_new['poa_monthly']]
    del [ac_monthly_new['solrad_monthly']]
    del [ac_monthly_new['dc_monthly']]
    del [ac_monthly_new['solrad_annual']]
    del [ac_monthly_new['capacity_factor']]
    ac_monthly_new = str(ac_monthly_new)
    ac_monthly_new = ac_monthly_new[14:]

    ch = '.'
    try:
        # Remove all characters after the character '.' from string
        ac_monthly_original = ac_monthly_original[0: ac_monthly_original.index(ch)]
    except ValueError:
        pass

    ch = '.'
    try:
        # Remove all characters after the character '.' from string
        ac_monthly_new = ac_monthly_new[0: ac_monthly_new.index(ch)]
    except ValueError:
        pass

    # ^^^^ finished parsing / quite the mess but works great, on to calculate percent difference between calls

    ac_monthly_original = int(ac_monthly_original)
    ac_monthly_new = int(ac_monthly_new)

    difference = ac_monthly_original - ac_monthly_new
    total = difference / ac_monthly_original

    if total <= 0.1:
        total = str(total)
        total = total[2:]
        total = total[:2]
        total = total[1:]
    else:
        total = str(total)
        total = total[2:]
        total = total[:2]

    total = str(total)
    total = total + "%"

    # calculations for ten percent docx sheet finished, some parsing of useless string data ******************

    # setting variables and values for ten percent docx and finishing out with a final print *****************
    # This is here to fix the TEN_PERCENT letter back to original formatting
    Old_tilt = str(tab_lookup.acell("M2").value)
    Old_azimuth = str(tab_lookup.acell("M3").value)
    New_tilt = str(tab_lookup.acell("N2").value)
    New_azimuth = str(tab_lookup.acell("N3").value)
    line_break = "________________________________________________"

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': Hoa_name, 'date': Date, 'name': name,
               'quantity': Quantity, 'old_direction': Old_direction, 'quantity2': Quantity_2, 'state': state,
               'old_azimuth': Old_azimuth, 'old_tilt': Old_tilt, 'new_direction': New_direction,
               'new_azimuth': New_azimuth, 'new_tilt': New_tilt, 'mod_watt': Mod_watt, 'percent': total,
               'ac_monthly_original': ac_monthly_original, 'ac_monthly_new': ac_monthly_new}

    doc.render(context)
    doc.save(name + " Ten Percent Letter Array 1.docx")
    print("Ten Percent Letter finished...")

def arrayTwo():

    Quantity = str(tab_lookup.acell("M12").value)
    Quantity_2 = str(tab_lookup.acell("N12").value)
    Old_tilt = str(tab_lookup.acell("M9").value)
    Old_azimuth = str(tab_lookup.acell("M10").value)
    Old_direction = str(tab_lookup.acell("M13").value)
    New_direction = str(tab_lookup.acell("N13").value)
    New_tilt = str(tab_lookup.acell("N9").value)
    New_azimuth = str(tab_lookup.acell("N10").value)
    Mod_watt = str(tab_lookup.acell("C10").value)
    Address = str(tab_lookup.acell("B13:C13").value)
    Losses_o = str(tab_lookup.acell("M11").value)
    Losses_n = str(tab_lookup.acell("N11").value)
    Module_type = "1"
    Array_type = "1"
    System_capacity = 1
    System_capacity_2 = 1
    # ^^^ Google sheet values - checked and stringed ready to pass into docxtpl and calculations ^^^ *********

    # Calculating System_capacity ****************************************************************************
    Quantity = int(Quantity)
    Quantity_2 = int(Quantity_2)

    if Mod_watt == "SPR-M435-H-AC":
        System_capacity = Quantity * .435
        System_capacity_2 = Quantity_2 * .435
    elif Mod_watt == "SPR-M425-H-AC":
        System_capacity = Quantity * .425
        System_capacity_2 = Quantity_2 * .425
    elif Mod_watt == "SPR-A420-AC":
        System_capacity = Quantity * .420
        System_capacity_2 = Quantity_2 * .420
    elif Mod_watt == "SPR-A415-AC":
        System_capacity = Quantity * .415
        System_capacity_2 = Quantity_2 * .415
    elif Mod_watt == "SPR-A410-AC":
        System_capacity = Quantity * .410
        System_capacity_2 = Quantity_2 * .410
    elif Mod_watt == "JKM410M-72HL-V G2 410W":
        System_capacity = Quantity * .410
        System_capacity_2 = Quantity_2 * .410
    elif Mod_watt == "SPR-A400-BLK-AC":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-A400-BLK":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-A400-AC":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-U400-BLK":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-X22-370-AC":
        System_capacity = Quantity * .370
        System_capacity_2 = Quantity_2 * .370
    elif Mod_watt == "SPR-X22-360-AC":
        System_capacity = Quantity * .360
        System_capacity_2 = Quantity_2 * .360
    elif Mod_watt == "SPR-X22-360":
        System_capacity = Quantity * .360
        System_capacity_2 = Quantity_2 * .360
    elif Mod_watt == "SPR-X21-350-BLK-AC":
        System_capacity = Quantity * .350
        System_capacity_2 = Quantity_2 * .350
    elif Mod_watt == "SPR-E20-327-AC":
        System_capacity = Quantity * .327
        System_capacity_2 = Quantity_2 * .327
    elif Mod_watt == "SPR-E20-327":
        System_capacity = Quantity * .327
        System_capacity_2 = Quantity_2 * .327
    elif Mod_watt == "SPR-E19-320-AC":
        System_capacity = Quantity * .320
        System_capacity_2 = Quantity_2 * .320
    else:
        pass

    System_capacity = str(System_capacity)
    System_capacity_2 = str(System_capacity_2)

    # *********************************************************************************************************

    # setting string variables for NREL/PVWATTS parameters ****************************************************

    Address = Address.replace(" ", "%20").strip()
    Address = ("address=" + Address + "&")
    New_tilt = ("tilt=" + New_tilt + "&")
    New_azimuth = ("azimuth=" + New_azimuth + "&")
    Old_tilt = ("tilt=" + Old_tilt + "&")
    Old_azimuth = ("azimuth=" + Old_azimuth + "&")
    Losses_o = ("losses=" + Losses_o + "&")
    Losses_n = ("losses=" + Losses_n + "&")
    Module_type = ("module_type=" + Module_type + "&")
    Array_type = ("array_type=" + Array_type + "&")
    System_capacity = ("system_capacity=" + System_capacity + "&")
    System_capacity_2 = ("system_capacity=" + System_capacity_2 + "&")

    API_PARAM = "&api_key="
    OLD_QUERY = Address + Old_tilt + Old_azimuth + Losses_o + Module_type + Array_type + System_capacity + API_PARAM + creds.API_KEY
    NEW_QUERY = Address + New_tilt + New_azimuth + Losses_n + Module_type + Array_type + System_capacity_2 + API_PARAM + creds.API_KEY

    # parameters set for NREL/PVW API connection, preparing to make get call & parse data after 200 response *

    # preforming first requests to NREL/PVW API -- Looks a little jank but it works **************************
    response = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + API_PARAM + creds.API_KEY)
    BASE_URL = "https://developer.nrel.gov/api/pvwatts/v6.json?"

    JSON_LINK_ORIGINAL = (BASE_URL + OLD_QUERY)
    JSON_LINK_NEW = (BASE_URL + NEW_QUERY)
    data_original = requests.get(JSON_LINK_ORIGINAL)
    data_new = requests.get(JSON_LINK_NEW)

    print(JSON_LINK_ORIGINAL)
    print(JSON_LINK_NEW)
    print(data_original.status_code)
    print(data_new.status_code)

    # finished api request should have 200 response in console / printed *************************************

    # had a lot of trouble with the json formatting from NREL/PVW had to parse this way to get annual total **
    content = data_original.text
    data_original = json.loads(content)
    content = data_new.text
    data_new = json.loads(content)

    ac_monthly_original = data_original.get('outputs')
    ac_monthly_new = data_new.get('outputs')
    dict.items(ac_monthly_original)
    dict.items(ac_monthly_new)

    del [ac_monthly_original['ac_monthly']]
    del [ac_monthly_original['poa_monthly']]
    del [ac_monthly_original['solrad_monthly']]
    del [ac_monthly_original['dc_monthly']]
    del [ac_monthly_original['solrad_annual']]
    del [ac_monthly_original['capacity_factor']]
    ac_monthly_original = str(ac_monthly_original)
    ac_monthly_original = ac_monthly_original[14:]

    del [ac_monthly_new['ac_monthly']]
    del [ac_monthly_new['poa_monthly']]
    del [ac_monthly_new['solrad_monthly']]
    del [ac_monthly_new['dc_monthly']]
    del [ac_monthly_new['solrad_annual']]
    del [ac_monthly_new['capacity_factor']]
    ac_monthly_new = str(ac_monthly_new)
    ac_monthly_new = ac_monthly_new[14:]

    ch = '.'
    try:
        # Remove all characters after the character '.' from string
        ac_monthly_original = ac_monthly_original[0: ac_monthly_original.index(ch)]
    except ValueError:
        pass

    ch = '.'
    try:
        # Remove all characters after the character '.' from string
        ac_monthly_new = ac_monthly_new[0: ac_monthly_new.index(ch)]
    except ValueError:
        pass

    # ^^^^ finished parsing / quite the mess but works great, on to calculate percent difference between calls

    ac_monthly_original = int(ac_monthly_original)
    ac_monthly_new = int(ac_monthly_new)

    difference = ac_monthly_original - ac_monthly_new
    total = difference / ac_monthly_original

    if total <= 0.1:
        total = str(total)
        total = total[2:]
        total = total[:2]
        total = total[1:]
    else:
        total = str(total)
        total = total[2:]
        total = total[:2]

    total = str(total)
    total = total + "%"

    # calculations for ten percent docx sheet finished, some parsing of useless string data ******************

    # setting variables and values for ten percent docx and finishing out with a final print *****************
    # This is here to fix the TEN_PERCENT letter back to original formatting
    Old_tilt = str(tab_lookup.acell("M9").value)
    Old_azimuth = str(tab_lookup.acell("M10").value)
    New_tilt = str(tab_lookup.acell("N9").value)
    New_azimuth = str(tab_lookup.acell("N10").value)
    line_break = "________________________________________________"

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': Hoa_name, 'date': Date, 'name': name,
               'quantity': Quantity, 'old_direction': Old_direction, 'quantity2': Quantity_2, 'state': state,
               'old_azimuth': Old_azimuth, 'old_tilt': Old_tilt, 'new_direction': New_direction,
               'new_azimuth': New_azimuth, 'new_tilt': New_tilt, 'mod_watt': Mod_watt, 'percent': total,
               'ac_monthly_original': ac_monthly_original, 'ac_monthly_new': ac_monthly_new}

    doc.render(context)
    doc.save(name + " Ten Percent Letter Array 2.docx")
    print("Ten Percent Letter finished...")

def arrayThree():

    Quantity = str(tab_lookup.acell("M19").value)
    Quantity_2 = str(tab_lookup.acell("N19").value)
    Old_tilt = str(tab_lookup.acell("M16").value)
    Old_azimuth = str(tab_lookup.acell("M17").value)
    Old_direction = str(tab_lookup.acell("M20").value)
    New_direction = str(tab_lookup.acell("N20").value)
    New_tilt = str(tab_lookup.acell("N16").value)
    New_azimuth = str(tab_lookup.acell("N17").value)
    Mod_watt = str(tab_lookup.acell("C10").value)
    Address = str(tab_lookup.acell("B13:C13").value)
    Losses_o = str(tab_lookup.acell("M18").value)
    Losses_n = str(tab_lookup.acell("N18").value)
    Module_type = "1"
    Array_type = "1"
    System_capacity = 1
    System_capacity_2 = 1
    # ^^^ Google sheet values - checked and stringed ready to pass into docxtpl and calculations ^^^ *********

    # Calculating System_capacity ****************************************************************************
    Quantity = int(Quantity)
    Quantity_2 = int(Quantity_2)

    if Mod_watt == "SPR-M435-H-AC":
        System_capacity = Quantity * .435
        System_capacity_2 = Quantity_2 * .435
    elif Mod_watt == "SPR-M425-H-AC":
        System_capacity = Quantity * .425
        System_capacity_2 = Quantity_2 * .425
    elif Mod_watt == "SPR-A420-AC":
        System_capacity = Quantity * .420
        System_capacity_2 = Quantity_2 * .420
    elif Mod_watt == "SPR-A415-AC":
        System_capacity = Quantity * .415
        System_capacity_2 = Quantity_2 * .415
    elif Mod_watt == "SPR-A410-AC":
        System_capacity = Quantity * .410
        System_capacity_2 = Quantity_2 * .410
    elif Mod_watt == "JKM410M-72HL-V G2 410W":
        System_capacity = Quantity * .410
        System_capacity_2 = Quantity_2 * .410
    elif Mod_watt == "SPR-A400-BLK-AC":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-A400-BLK":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-A400-AC":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-U400-BLK":
        System_capacity = Quantity * .400
        System_capacity_2 = Quantity_2 * .400
    elif Mod_watt == "SPR-X22-370-AC":
        System_capacity = Quantity * .370
        System_capacity_2 = Quantity_2 * .370
    elif Mod_watt == "SPR-X22-360-AC":
        System_capacity = Quantity * .360
        System_capacity_2 = Quantity_2 * .360
    elif Mod_watt == "SPR-X22-360":
        System_capacity = Quantity * .360
        System_capacity_2 = Quantity_2 * .360
    elif Mod_watt == "SPR-X21-350-BLK-AC":
        System_capacity = Quantity * .350
        System_capacity_2 = Quantity_2 * .350
    elif Mod_watt == "SPR-E20-327-AC":
        System_capacity = Quantity * .327
        System_capacity_2 = Quantity_2 * .327
    elif Mod_watt == "SPR-E20-327":
        System_capacity = Quantity * .327
        System_capacity_2 = Quantity_2 * .327
    elif Mod_watt == "SPR-E19-320-AC":
        System_capacity = Quantity * .320
        System_capacity_2 = Quantity_2 * .320
    else:
        pass

    System_capacity = str(System_capacity)
    System_capacity_2 = str(System_capacity_2)

    # *********************************************************************************************************

    # setting string variables for NREL/PVWATTS parameters ****************************************************

    Address = Address.replace(" ", "%20").strip()
    Address = ("address=" + Address + "&")
    New_tilt = ("tilt=" + New_tilt + "&")
    New_azimuth = ("azimuth=" + New_azimuth + "&")
    Old_tilt = ("tilt=" + Old_tilt + "&")
    Old_azimuth = ("azimuth=" + Old_azimuth + "&")
    Losses_o = ("losses=" + Losses_o + "&")
    Losses_n = ("losses=" + Losses_n + "&")
    Module_type = ("module_type=" + Module_type + "&")
    Array_type = ("array_type=" + Array_type + "&")
    System_capacity = ("system_capacity=" + System_capacity + "&")
    System_capacity_2 = ("system_capacity=" + System_capacity_2 + "&")

    API_PARAM = "&api_key="
    OLD_QUERY = Address + Old_tilt + Old_azimuth + Losses_o + Module_type + Array_type + System_capacity + API_PARAM + creds.API_KEY
    NEW_QUERY = Address + New_tilt + New_azimuth + Losses_n + Module_type + Array_type + System_capacity_2 + API_PARAM + creds.API_KEY

    # parameters set for NREL/PVW API connection, preparing to make get call & parse data after 200 response *

    # preforming first requests to NREL/PVW API -- Looks a little jank but it works **************************
    response = requests.get("https://developer.nrel.gov/api/pvwatts/v6.json?" + API_PARAM + creds.API_KEY)
    BASE_URL = "https://developer.nrel.gov/api/pvwatts/v6.json?"

    JSON_LINK_ORIGINAL = (BASE_URL + OLD_QUERY)
    JSON_LINK_NEW = (BASE_URL + NEW_QUERY)
    data_original = requests.get(JSON_LINK_ORIGINAL)
    data_new = requests.get(JSON_LINK_NEW)

    print(JSON_LINK_ORIGINAL)
    print(JSON_LINK_NEW)
    print(data_original.status_code)
    print(data_new.status_code)

    # finished api request should have 200 response in console / printed *************************************

    # had a lot of trouble with the json formatting from NREL/PVW had to parse this way to get annual total **
    content = data_original.text
    data_original = json.loads(content)
    content = data_new.text
    data_new = json.loads(content)

    ac_monthly_original = data_original.get('outputs')
    ac_monthly_new = data_new.get('outputs')
    dict.items(ac_monthly_original)
    dict.items(ac_monthly_new)

    del [ac_monthly_original['ac_monthly']]
    del [ac_monthly_original['poa_monthly']]
    del [ac_monthly_original['solrad_monthly']]
    del [ac_monthly_original['dc_monthly']]
    del [ac_monthly_original['solrad_annual']]
    del [ac_monthly_original['capacity_factor']]
    ac_monthly_original = str(ac_monthly_original)
    ac_monthly_original = ac_monthly_original[14:]

    del [ac_monthly_new['ac_monthly']]
    del [ac_monthly_new['poa_monthly']]
    del [ac_monthly_new['solrad_monthly']]
    del [ac_monthly_new['dc_monthly']]
    del [ac_monthly_new['solrad_annual']]
    del [ac_monthly_new['capacity_factor']]
    ac_monthly_new = str(ac_monthly_new)
    ac_monthly_new = ac_monthly_new[14:]

    ch = '.'
    try:
        # Remove all characters after the character '.' from string
        ac_monthly_original = ac_monthly_original[0: ac_monthly_original.index(ch)]
    except ValueError:
        pass

    ch = '.'
    try:
        # Remove all characters after the character '.' from string
        ac_monthly_new = ac_monthly_new[0: ac_monthly_new.index(ch)]
    except ValueError:
        pass

    # ^^^^ finished parsing / quite the mess but works great, on to calculate percent difference between calls

    ac_monthly_original = int(ac_monthly_original)
    ac_monthly_new = int(ac_monthly_new)

    difference = ac_monthly_original - ac_monthly_new
    total = difference / ac_monthly_original

    if total <= 0.1:
        total = str(total)
        total = total[2:]
        total = total[:2]
        total = total[1:]
    else:
        total = str(total)
        total = total[2:]
        total = total[:2]

    total = str(total)
    total = total + "%"

    # calculations for ten percent docx sheet finished, some parsing of useless string data ******************

    # setting variables and values for ten percent docx and finishing out with a final print *****************
    # This is here to fix the TEN_PERCENT letter back to original formatting
    Old_tilt = str(tab_lookup.acell("M16").value)
    Old_azimuth = str(tab_lookup.acell("M17").value)
    New_tilt = str(tab_lookup.acell("N16").value)
    New_azimuth = str(tab_lookup.acell("N17").value)
    line_break = "________________________________________________"

    doc = DocxTemplate("TEN_PERCENT_V5.docx")
    context = {'hoa_name': Hoa_name, 'date': Date, 'name': name,
               'quantity': Quantity, 'old_direction': Old_direction, 'quantity2': Quantity_2, 'state': state,
               'old_azimuth': Old_azimuth, 'old_tilt': Old_tilt, 'new_direction': New_direction,
               'new_azimuth': New_azimuth, 'new_tilt': New_tilt, 'mod_watt': Mod_watt, 'percent': total,
               'ac_monthly_original': ac_monthly_original, 'ac_monthly_new': ac_monthly_new}

    doc.render(context)
    doc.save(name + " Ten Percent Letter Array 3.docx")
    print("Ten Percent Letter finished...")

def main():

    if arrayCount == 1:
        arrayOne()
    elif arrayCount == 2:
        arrayOne()
        arrayTwo()
    elif arrayCount == 3:
        arrayOne()
        arrayTwo()
        arrayThree()
    elif arrayCount == 4:
        pass
    else:
        exit()

if __name__ == '__main__':
    main()