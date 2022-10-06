import os
import copy
import magic
import shutil
import string
import hashlib
import xml.etree.ElementTree as ET
from uuid import uuid1
from datetime import datetime
from subprocess import run, DEVNULL
from tkinter import Tk, messagebox, filedialog, Label, Button, Entry, ttk, StringVar

municipality_list = sorted(["5041 Snåsa Kommune", "5057 Ørland Kommune", "5059 Orkland Kommune", "5034 Meråker Kommune", "5037 Levanger Kommune", "5025 Røros Kommune", "5016 Agdenes Kommune", "5012 Snillfjord Kommune", "5036 Frosta Kommune", "5023 Meldal Kommune", "5044 Namsskogan Kommune", "5043 Røyrvik Kommune", "5011 Hemne Kommune", "5032 Selbu Kommune", "5035 Stjørdal Kommune", "5046 Høylandet Kommune", "5042 Lierne Kommune", "5045 Grong Kommune","5049 Flatanger Kommune","5014 Frøya Kommune","5055 Heim Kommune","5013 Hitra Kommune","5026 Holtålen Kommune","5053 Inderøy Kommune","5054 Indre Fosen Kommune","5031 Malvik Kommune","5028 Melhus Kommune","5027 Midtre Gauldal Kommune","5005 Namsos Kommune","5060 Nærøysund Kommune","5021 Oppdal Kommune","3430 Os Kommune","5047 Overhalla Kommune","5020 Osen Kommune","5022 Rennebu Kommune","5029 Skaun Kommune","5006 Steinkjer Kommune","5033 Tydal Kommune","5038 Verdal Kommune","5058 Åfjord Kommune"], key=lambda x: x.split(" ")[1])
system_list = sorted(["ESA", "Visma Velferd", "Visma Familia", "WinMed Helse", "Ephorte", "Visma Flyt Skole", "Visma Profil", "SystemX", "P360", "Digora", "Oppad", "CGM Helsestasjon", "Visma Flyt Sampro", "Gerica", "Socio"])
info_dict = {}

#Function to browse computer for files
def browse_files(label):
    file = filedialog.askdirectory(initialdir="./", title="Choose a folder whose content should be packaged")
    label.configure(text=file)

def configure_sip_log(log_path, id, create_date):
    print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Configuring sip log.xml...')
    ET.register_namespace('premis', "http://arkivverket.no/standarder/PREMIS")
    logfile = ET.parse(log_path)
    xml_root = logfile.getroot()
    xml_root.attrib["xmlns:xlink"] = "http://www.w3.org/1999/xlink"
    
    xml_root[0][0][1].text = id
    xml_root[0][3][1].text = create_date
    xml_root[0][4][1].text = archivist_org_combo.get()
    xml_root[0][5][1].text = label_entry.get()
    xml_root[1][0][1].text = str(uuid1())
    xml_root[1][2].text = create_date
    xml_root[1][6][1].text = id
    
    logfile.write(log_path, encoding="UTF-8", xml_declaration=True, short_empty_elements=False)


def configure_sip_premis(premis_path, id, tarfolder):
    print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Configuring sip premis.xml...')
    premisfile = ET.parse(premis_path)
    xml_root = premisfile.getroot()
    xml_root.attrib["xmlns:xlink"] = "http://www.w3.org/1999/xlink"
    xml_root[0][0][1].text = id
    for i in range(1, 3):
        xml_root[i][0][1].text = id + xml_root[i][0][1].text
        xml_root[i][2][0][1].text = id
        xml_root[i][3][2][1].text = id

    for root, _, files in os.walk(tarfolder):
        for file in files:
            tmp_path = root.split("/")[4].replace("\\", "/") + "/" + file
            if tmp_path != f'{id}/administrative_metadata/DIAS_PREMIS.xsd' and tmp_path != f'{id}/mets.xsd' and tmp_path != f'{id}/mets.xml' and tmp_path != f'{id}/administrative_metadata/premis.xml':
                xml_root.insert(len(xml_root)-1, copy.deepcopy(xml_root[1]))
                xml_root[-2][0][1].text = tmp_path
                if tmp_path not in info_dict.keys():
                    sha = hashlib.sha256()
                    with open(os.path.join(root, file), "rb") as f:
                        tmp_data = f.read(4000000)
                        tmp_magic = magic.from_buffer(tmp_data, mime=True)
                        while len(tmp_data)>0:
                            sha.update(tmp_data)
                            tmp_data = f.read(4000000)
                    info_dict[tmp_path] = [sha.hexdigest(), tmp_magic]
                xml_root[-2][1][1][1].text = info_dict[tmp_path][0]
                xml_root[-2][1][2].text = str(os.stat(os.path.join(root, file)).st_size)
                xml_root[-2][1][3][0][0].text = os.path.splitext(file)[1][1:]

    ET.indent(premisfile, '  ')
    premisfile.write(premis_path, encoding="UTF-8", xml_declaration=True, short_empty_elements=False)

def configure_sip_mets(mets_path, id, tarfolder, creation_date, premis_path):
    print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Configuring sip mets.xml...')
    ET.register_namespace('mets', "http://www.loc.gov/METS/")
    ET.register_namespace('xlink', "http://www.w3.org/1999/xlink")
    metsfile = ET.parse(mets_path)
    xml_root = metsfile.getroot()
    xml_root.attrib.update({"OBJID": f'UUID:{id}', "ID": f'ID{uuid1()}', "LABEL": label_entry.get()})

    xml_root[0].attrib["CREATEDATE"] = creation_date
    xml_root[0][0][0].text = archivist_org_combo.get() #Archivist
    xml_root[0][1][0].text = system_combo.get() #System name
    xml_root[0][2][0].text = system_ver_entry.get() #System version
    xml_root[0][3][0].text = type_combo.get() #Type
    xml_root[0][4][0].text = creator_entry.get() #Creator
    xml_root[0][5][0].text = producer_org_entry.get() #Producer org
    xml_root[0][6][0].text = producer_pers_entry.get() #Producer person
    xml_root[0][7][0].text = producer_software_entry.get() #Producer software
    xml_root[0][8][0].text = submitter_org_combo.get() #Submitter org
    xml_root[0][9][0].text = submitter_pers_entry.get() #Submitter person
    xml_root[0][10][0].text = owner_org_combo.get() #Owner
    xml_root[0][11][0].text = preserver_entry.get() #Preserver
    xml_root[0][12].text = submission_entry.get() #ID
    xml_root[0][13].text = period_start_entry.get() #Start
    xml_root[0][14].text = period_end_entry.get() #End

    xml_root[1][0][0].attrib.update({"SIZE": str(os.stat(premis_path).st_size), "ID": f'ID{uuid1()}', "CREATED": datetime.fromtimestamp(os.path.getmtime(premis_path)).strftime("%Y-%m-%dT%H:%M:%S+02:00")})
    xml_root[3][0][0][0].attrib["FILEID"] = xml_root[1][0][0].attrib["ID"]
    with open(premis_path, "rb") as f:
        tmp_data = f.read()
    xml_root[1][0][0].attrib["CHECKSUM"] = hashlib.sha256(tmp_data).hexdigest()

    for root, _, files in os.walk(tarfolder):
        for file in files:
            tmp_path = "file:" + (root.replace("\\", "/") + "/" + file).removeprefix(f'{tarfolder}/')
            if tmp_path != "file:administrative_metadata/premis.xml" and tmp_path != "file:mets.xml":
                if tmp_path == "file:administrative_metadata/DIAS_PREMIS.xsd":
                    xml_root[2][0][0].attrib.update({"ID": f'ID{uuid1()}', "CREATED": datetime.fromtimestamp(os.path.getmtime(os.path.join(root,file))).strftime("%Y-%m-%dT%H:%M:%S+02:00")})
                    xml_root[3][0][1][0].attrib["FILEID"] = xml_root[2][0][0].attrib["ID"]
                elif tmp_path == "file:mets.xsd":
                    xml_root[2][0][1].attrib.update({"ID": f'ID{uuid1()}', "CREATED": datetime.fromtimestamp(os.path.getmtime(os.path.join(root,file))).strftime("%Y-%m-%dT%H:%M:%S+02:00")})
                    xml_root[3][0][1][1].attrib["FILEID"] = xml_root[2][0][1].attrib["ID"]
                else:
                    xml_root[2][0].append(copy.deepcopy(xml_root[2][0][0]))
                    xml_root[3][0][1].append(copy.deepcopy(xml_root[3][0][1][0]))
                    xml_root[2][0][-1].attrib.update({"SIZE": str(os.stat(os.path.join(root,file)).st_size), "ID": f'ID{uuid1()}', "CHECKSUM": info_dict[f'{id}/{tmp_path[5:]}'][0], "MIMETYPE": info_dict[f'{id}/{tmp_path[5:]}'][1], "CREATED": datetime.fromtimestamp(os.path.getmtime(os.path.join(root,file))).strftime("%Y-%m-%dT%H:%M:%S+02:00")})
                    xml_root[2][0][-1][0].attrib["{http://www.w3.org/1999/xlink}href"] = tmp_path
                    xml_root[3][0][1][-1].attrib["FILEID"] = xml_root[2][0][-1].attrib["ID"]

    ET.indent(metsfile, '  ')
    metsfile.write(mets_path, encoding="UTF-8", xml_declaration=True, short_empty_elements=True)

def configure_sip_info(info_path, tar_path, id, creation_date):
    print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Configuring sip info.xml...')
    infofile = ET.parse(info_path)
    xml_root = infofile.getroot()
    xml_root.attrib.update({"OBJID": f'UUID:{id}', "ID": f'ID{uuid1()}', "LABEL": label_entry.get()})
    
    xml_root[0].attrib["CREATEDATE"] = creation_date
    xml_root[0][0][0].text = archivist_org_combo.get() #Archivist
    xml_root[0][1][0].text = system_combo.get() #System name
    xml_root[0][2][0].text = system_ver_entry.get() #System version
    xml_root[0][3][0].text = type_combo.get() #Type
    xml_root[0][4][0].text = creator_entry.get() #Creator
    xml_root[0][5][0].text = producer_org_entry.get() #Producer org
    xml_root[0][6][0].text = producer_pers_entry.get() #Producer person
    xml_root[0][7][0].text = producer_software_entry.get() #Producer software
    xml_root[0][8][0].text = submitter_org_combo.get() #Submitter org
    xml_root[0][9][0].text = submitter_pers_entry.get() #Submitter person
    xml_root[0][10][0].text = owner_org_combo.get() #Owner
    xml_root[0][11][0].text = preserver_entry.get() #Preserver
    xml_root[0][12].text = submission_entry.get() #ID
    xml_root[0][13].text = period_start_entry.get() #Start
    xml_root[0][14].text = period_end_entry.get() #End

    xml_root[1][0][0].attrib.update({"SIZE": str(os.stat(tar_path).st_size), "ID": f'ID{uuid1()}', "CREATED": datetime.fromtimestamp(os.path.getmtime(tar_path)).strftime("%Y-%m-%dT%H:%M:%S+02:00")})
    xml_root[2][0][1][0].attrib["FILEID"] = xml_root[1][0][0].attrib["ID"]
    sha = hashlib.sha256()
    with open(tar_path, "rb") as f:
        tmp_data = f.read(4000000)
        while len(tmp_data)>0:
            sha.update(tmp_data)
            tmp_data = f.read(4000000)
    xml_root[1][0][0].attrib["CHECKSUM"] = sha.hexdigest()
    xml_root[1][0][0][0].attrib["{http://www.w3.org/1999/xlink}href"] = f'file:{os.path.basename(tar_path)}'

    ET.indent(infofile, '  ')
    infofile.write(info_path, encoding="UTF-8", xml_declaration=True, short_empty_elements=True)

def configure_aic_log(log_path, aic_id, sip_id, create_date):
    print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Configuring aic log.xml...')
    logfile = ET.parse(log_path)
    xml_root = logfile.getroot()
    xml_root.attrib["xmlns:xlink"] = "http://www.w3.org/1999/xlink"
    
    xml_root[0][0][1].text = sip_id
    xml_root[0][2][1].text = aic_id
    xml_root[0][3][1].text = create_date
    xml_root[0][4][1].text = archivist_org_combo.get()
    xml_root[0][5][1].text = label_entry.get()
    xml_root[0][9][2][1].text = aic_id
    xml_root[1][0][1].text = str(uuid1())
    xml_root[1][2].text = create_date
    xml_root[1][6][1].text = sip_id
    
    logfile.write(log_path, encoding="UTF-8", xml_declaration=True, short_empty_elements=False)

#Main function
def main_func():
    if content_path_label.cget("text"):
        print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Starting process...')
        sip_id = uuid1()

        template_folder = "./template"
        output_folder = f'./out'
        tarfile = f'{output_folder}/{sip_id}/content/{sip_id}'
        
        shutil.copytree(template_folder, output_folder, copy_function=shutil.copy)
        shutil.copytree(os.path.abspath(content_path_label.cget("text")), f'{output_folder}/top/content/bottom/content', copy_function=shutil.copy, dirs_exist_ok=True)
        if descriptive_path_label.cget("text"):
            shutil.copytree(os.path.abspath(descriptive_path_label.cget("text")), f'{output_folder}/top/content/bottom/descriptive_metadata', copy_function=shutil.copy, dirs_exist_ok=True)
        if administrative_path_label.cget("text"):
            shutil.copytree(os.path.abspath(administrative_path_label.cget("text")), f'{output_folder}/top/content/bottom/administrative_metadata', copy_function=shutil.copy, dirs_exist_ok=True)
        
        os.rename(f'{output_folder}/top', f'{output_folder}/{sip_id}')
        os.rename(f'{output_folder}/{sip_id}/content/bottom', tarfile)

        #Zone 1
        configure_sip_log(f'{tarfile}/log.xml', str(sip_id), datetime.now().strftime("%Y-%m-%dT%H:%M:%S+02:00"))
        configure_sip_premis(f'{tarfile}/administrative_metadata/premis.xml', str(sip_id), tarfile)
        configure_sip_mets(f'{tarfile}/mets.xml', str(sip_id), tarfile, datetime.now().strftime("%Y-%m-%dT%H:%M:%S+02:00"), f'{tarfile}/administrative_metadata/premis.xml')
        run([f'{os.environ["PROGRAMFILES"]}\\7-Zip\\7z.exe', 'a', f'{tarfile}.tar', tarfile, '-aou', '-sdel'], stdout=DEVNULL, stderr=DEVNULL)
        configure_sip_info(f'{output_folder}/info.xml', f'{tarfile}.tar', str(sip_id), datetime.now().strftime("%Y-%m-%dT%H:%M:%S+02:00"))

        #Zone 2
        aic_id = uuid1()
        os.rename(output_folder, f'./{aic_id}')
        configure_aic_log(f'./{aic_id}/{sip_id}/log.xml', str(aic_id), str(sip_id), datetime.now().strftime("%Y-%m-%dT%H:%M:%S+02:00"))

        messagebox.showinfo("Result", "Process complete!")

#Validates all input given to the entry/combo-boxes in the gui. Returns True if the input is valid
def validation_func(val, rem):
    if rem == "1":
        for char in val.lower():
            if char not in f'{string.printable}æøå' or char in "\'\"<>&":
                return False
    return True

#A helper function which makes the relevant combo-box only list results relevant to what is already written in it
def combo_helper(element, list):
    element['values'] = [i for i in list if element.get().lower() in i.lower()]

#Initializes the main window
window = Tk()
window.title("Archive Package Creator")
window.geometry('{width}x{height}+{pos_right}+{pos_down}'.format(width=(window.winfo_screenwidth() // 2)+(window.winfo_screenwidth() // 6), height=(window.winfo_screenheight() // 2)+(window.winfo_screenheight() // 6), pos_right=(window.winfo_screenwidth() // 4)-(window.winfo_screenwidth() // 12), pos_down=(window.winfo_screenheight() // 4)-(window.winfo_screenheight() // 12)))
window.config(background="teal")

#Configures columns and rows in the main window's grid
for i in range(20):
    window.columnconfigure(i, weight=1, pad=0)
    window.rowconfigure(i, weight=1, pad=0)

#Creates labels, combo-boxes, entrys and some buttons with self-explanatory names
content_path_label = Label(window, text="", relief="solid")
content_path_button = Button(window, text="Browse Content", command=lambda: browse_files(content_path_label))
descriptive_path_label = Label(window, text="", relief="solid")
descriptive_path_button = Button(window, text="Browse Descriptive Metadata (optional)", command=lambda: browse_files(descriptive_path_label))
administrative_path_label = Label(window, text="", relief="solid")
administrative_path_button = Button(window, text="Browse Administrative Metadata (optional)", command=lambda: browse_files(administrative_path_label))
system_label = Label(window, text="System:", background="black", foreground="white")
system_combo = ttk.Combobox(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"), postcommand=lambda: combo_helper(system_combo,system_list))
system_ver_label = Label(window, text="System Version:", background="black", foreground="white")
system_ver_entry = Entry(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
submission_label = Label(window, text="Submission Agreement:", background="black", foreground="white")
submission_entry = Entry(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
archivist_org_label = Label(window, text="Archivist Organization:", background="black", foreground="white")
archivist_org_combo = ttk.Combobox(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"), postcommand=lambda: combo_helper(archivist_org_combo,municipality_list))
label_label = Label(window, text="Label:", background="black", foreground="white")
label_entry = Entry(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
type_label = Label(window, text="Archivist System Type:", background="black", foreground="white")
type_combo = ttk.Combobox(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"), postcommand=lambda: combo_helper(type_combo,["SIARD", "NOARK-5", "Postjournaler"]))
owner_org_label = Label(window, text="Owner Organization:", background="black", foreground="white")
owner_org_combo = ttk.Combobox(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"), postcommand=lambda: combo_helper(owner_org_combo,municipality_list))
producer_org_label = Label(window, text="Producer Organization:", background="black", foreground="white")
producer_org_entry = Entry(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
producer_pers_label = Label(window, text="Producer Person:", background="black", foreground="white")
producer_pers_entry = Entry(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
producer_software_label = Label(window, text="Producer Software:", background="black", foreground="white")
producer_software_entry = Entry(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
period_start_label = Label(window, text="Period Start:", background="black", foreground="white")
period_start_entry = Entry(window, textvariable=StringVar(window, "yyyy-mm-dd"), validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
period_end_label = Label(window, text="Period End:", background="black", foreground="white")
period_end_entry = Entry(window, textvariable=StringVar(window, "yyyy-mm-dd"), validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
submitter_org_label = Label(window, text="Submitter Organization:", background="black", foreground="white")
submitter_org_combo = ttk.Combobox(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"), postcommand=lambda: combo_helper(submitter_org_combo,municipality_list))
submitter_pers_label = Label(window, text="Submitter Person:", background="black", foreground="white")
submitter_pers_entry = Entry(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
creator_label = Label(window, text="Creator:", background="black", foreground="white")
creator_entry = Entry(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
preserver_label = Label(window, text="Preserver:", background="black", foreground="white")
preserver_entry = Entry(window, validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
button_execute = Button(window, text="Create Dias Package", command=main_func)

#Adds the widgets to the main window
content_path_label.grid(column=1, row=1, columnspan=13, sticky="NSEW")
content_path_button.grid(column=16, row=1, columnspan=3, sticky="NSEW")
descriptive_path_label.grid(column=1, row=2, columnspan=13, sticky="NSEW")
descriptive_path_button.grid(column=16, row=2, columnspan=3, sticky="NSEW")
administrative_path_label.grid(column=1, row=3, columnspan=13, sticky="NSEW")
administrative_path_button.grid(column=16, row=3, columnspan=3, sticky="NSEW")
system_label.grid(column=1, row=5, columnspan=2, sticky="NSEW")
system_combo.grid(column=3, row=5, columnspan=7, sticky="NSEW")
system_ver_label.grid(column=10, row=5, columnspan=2, sticky="NSEW")
system_ver_entry.grid(column=12, row=5, columnspan=7, sticky="NSEW")
submission_label.grid(column=1, row=6, columnspan=2, sticky="NSEW")
submission_entry.grid(column=3, row=6, columnspan=16, sticky="NSEW")
archivist_org_label.grid(column=1, row=7, columnspan=2, sticky="NSEW")
archivist_org_combo.grid(column=3, row=7, columnspan=16, sticky="NSEW")
label_label.grid(column=1, row=8, columnspan=2, sticky="NSEW")
label_entry.grid(column=3, row=8, columnspan=16, sticky="NSEW")
type_label.grid(column=1, row=9, columnspan=2, sticky="NSEW")
type_combo.grid(column=3, row=9, columnspan=16, sticky="NSEW")
owner_org_label.grid(column=1, row=10, columnspan=2, sticky="NSEW")
owner_org_combo.grid(column=3, row=10, columnspan=16, sticky="NSEW")
producer_org_label.grid(column=1, row=11, columnspan=2, sticky="NSEW")
producer_org_entry.grid(column=3, row=11, columnspan=16, sticky="NSEW")
producer_pers_label.grid(column=1, row=12, columnspan=2, sticky="NSEW")
producer_pers_entry.grid(column=3, row=12, columnspan=16, sticky="NSEW")
producer_software_label.grid(column=1, row=13, columnspan=2, sticky="NSEW")
producer_software_entry.grid(column=3, row=13, columnspan=16, sticky="NSEW")
period_start_label.grid(column=1, row=14, columnspan=2, sticky="NSEW")
period_start_entry.grid(column=3, row=14, columnspan=7, sticky="NSEW")
period_end_label.grid(column=10, row=14, columnspan=2, sticky="NSEW")
period_end_entry.grid(column=12, row=14, columnspan=7, sticky="NSEW")
submitter_org_label.grid(column=1, row=15, columnspan=2, sticky="NSEW")
submitter_org_combo.grid(column=3, row=15, columnspan=7, sticky="NSEW")
submitter_pers_label.grid(column=10, row=15, columnspan=2, sticky="NSEW")
submitter_pers_entry.grid(column=12, row=15, columnspan=7, sticky="NSEW")
creator_label.grid(column=1, row=16, columnspan=2, sticky="NSEW")
creator_entry.grid(column=3, row=16, columnspan=7, sticky="NSEW")
preserver_label.grid(column=10, row=16, columnspan=2, sticky="NSEW")
preserver_entry.grid(column=12, row=16, columnspan=7, sticky="NSEW")
button_execute.grid(column=1, row=18, columnspan=18, sticky="NSEW")

#Runs the main window and keeps it going
window.mainloop()
