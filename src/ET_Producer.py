import os
import sys
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

MUNICIPALITY_LIST = sorted(["5041 Snåsa Kommune", "5057 Ørland Kommune", "5059 Orkland Kommune", "5034 Meråker Kommune", "5037 Levanger Kommune", "5025 Røros Kommune", "5016 Agdenes Kommune", "5012 Snillfjord Kommune", "5036 Frosta Kommune", "5023 Meldal Kommune", "5044 Namsskogan Kommune", "5043 Røyrvik Kommune", "5011 Hemne Kommune", "5032 Selbu Kommune", "5035 Stjørdal Kommune", "5046 Høylandet Kommune", "5042 Lierne Kommune", "5045 Grong Kommune","5049 Flatanger Kommune","5014 Frøya Kommune","5055 Heim Kommune","5013 Hitra Kommune","5026 Holtålen Kommune","5053 Inderøy Kommune","5054 Indre Fosen Kommune","5031 Malvik Kommune","5028 Melhus Kommune","5027 Midtre Gauldal Kommune","5005 Namsos Kommune","5060 Nærøysund Kommune","5021 Oppdal Kommune","3430 Os Kommune","5047 Overhalla Kommune","5020 Osen Kommune","5022 Rennebu Kommune","5029 Skaun Kommune","5006 Steinkjer Kommune","5033 Tydal Kommune","5038 Verdal Kommune","5058 Åfjord Kommune"], key=lambda x: x.split(" ")[1])
SYSTEM_LIST = sorted(["ESA", "Visma Velferd", "Visma Familia", "WinMed Helse", "Ephorte", "Visma Flyt Skole", "Visma Profil", "SystemX", "P360", "Digora", "Oppad", "CGM Helsestasjon", "Visma Flyt Sampro", "Gerica", "Socio"])
INFO_DICT = {}

#Function to browse computer for files
def browse_files(label):
    file = filedialog.askdirectory(initialdir="./", title="Choose a folder whose content should be packaged")
    label.configure(text=file)

#Function to import metadata from chosen mets.xml file
def import_metadata(path):
    compare_dict = {frozenset(["ORGANIZATION","ARCHIVIST"]):[TEXT_LIST[3]], frozenset(["OTHER","SOFTWARE","ARCHIVIST"]):[TEXT_LIST[0],TEXT_LIST[1],TEXT_LIST[5]], frozenset(["ORGANIZATION","OTHER","PRODUCER"]):[TEXT_LIST[7]], frozenset(["INDIVIDUAL","OTHER","PRODUCER"]):[TEXT_LIST[8]], frozenset(["OTHER","SOFTWARE","OTHER","PRODUCER"]):[TEXT_LIST[9]], frozenset(["ORGANIZATION","OTHER","SUBMITTER"]):[TEXT_LIST[12]], frozenset(["INDIVIDUAL","OTHER","SUBMITTER"]):[TEXT_LIST[13]], frozenset(["ORGANIZATION","IPOWNER"]):[TEXT_LIST[6]], frozenset(["STARTDATE"]):[TEXT_LIST[10]], frozenset(["ENDDATE"]):[TEXT_LIST[11]]}
    if path:
        with open(path.name, "rb") as f:
            xml = f.read()
            xml = xml[xml.index(b'<?xml'):]
        xml_root = ET.fromstring(xml)
        if "http://www.loc.gov/METS/" in xml_root.tag:
            for i in xml_root[0]:
                if len(list(i)) and frozenset(i.attrib.values()) in compare_dict:
                    for j in range(min(len(list(i)), len(compare_dict[frozenset(i.attrib.values())]))):
                        if i[j].text:
                            compare_dict[frozenset(i.attrib.values())][j].set(i[j].text)
                elif i.text and frozenset(i.attrib.values()) in compare_dict:
                    compare_dict[frozenset(i.attrib.values())][0].set(i.text)
        else:
            messagebox.showerror("Error", "Unable to recognise file as a mets.xml file.")

def configure_sip_log(log_path, id, create_date):
    print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Configuring sip log.xml...')
    with open(log_path, "w", encoding="utf-8") as fo:
        string_log = f'<?xml version=\'1.0\' encoding=\'UTF-8\'?>\n<premis:premis xmlns:premis="http://arkivverket.no/standarder/PREMIS" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink" xsi:schemaLocation="http://arkivverket.no/standarder/PREMIS http://schema.arkivverket.no/PREMIS/v2.0/DIAS_PREMIS.xsd" version="2.0">\n  <premis:object xsi:type="premis:file">\n    <premis:objectIdentifier>\n      <premis:objectIdentifierType>NO/RA</premis:objectIdentifierType>\n      <premis:objectIdentifierValue>{id}</premis:objectIdentifierValue>\n    </premis:objectIdentifier>\n    <premis:preservationLevel>\n      <premis:preservationLevelValue>full</premis:preservationLevelValue>\n    </premis:preservationLevel>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>aic_object</premis:significantPropertiesType>\n      <premis:significantPropertiesValue></premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>createdate</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{create_date}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>archivist_organization</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{archivist_org_combo.get()}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>label</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{label_entry.get()}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>iptype</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>SIP</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:objectCharacteristics>\n      <premis:compositionLevel>0</premis:compositionLevel>\n      <premis:format>\n        <premis:formatDesignation>\n          <premis:formatName>tar</premis:formatName>\n        </premis:formatDesignation>\n      </premis:format>\n    </premis:objectCharacteristics>\n    <premis:storage>\n      <premis:storageMedium>Preservation platform ESSArch</premis:storageMedium>\n    </premis:storage>\n    <premis:relationship>\n      <premis:relationshipType>structural</premis:relationshipType>\n      <premis:relationshipSubType>is part of</premis:relationshipSubType>\n      <premis:relatedObjectIdentification>\n        <premis:relatedObjectIdentifierType>NO/RA</premis:relatedObjectIdentifierType>\n        <premis:relatedObjectIdentifierValue></premis:relatedObjectIdentifierValue>\n      </premis:relatedObjectIdentification>\n    </premis:relationship>\n  </premis:object>\n  <premis:event>\n    <premis:eventIdentifier>\n      <premis:eventIdentifierType>NO/RA</premis:eventIdentifierType>\n      <premis:eventIdentifierValue>{uuid1()}</premis:eventIdentifierValue>\n    </premis:eventIdentifier>\n    <premis:eventType>10000</premis:eventType>\n    <premis:eventDateTime>{create_date}</premis:eventDateTime>\n    <premis:eventDetail>Log circular created</premis:eventDetail>\n    <premis:eventOutcomeInformation>\n      <premis:eventOutcome>0</premis:eventOutcome>\n      <premis:eventOutcomeDetail>\n        <premis:eventOutcomeDetailNote>Success to create logfile</premis:eventOutcomeDetailNote>\n      </premis:eventOutcomeDetail>\n    </premis:eventOutcomeInformation>\n    <premis:linkingAgentIdentifier>\n      <premis:linkingAgentIdentifierType>NO/RA</premis:linkingAgentIdentifierType>\n      <premis:linkingAgentIdentifierValue>admin</premis:linkingAgentIdentifierValue>\n    </premis:linkingAgentIdentifier>\n    <premis:linkingObjectIdentifier>\n      <premis:linkingObjectIdentifierType>NO/RA</premis:linkingObjectIdentifierType>\n      <premis:linkingObjectIdentifierValue>{id}</premis:linkingObjectIdentifierValue>\n    </premis:linkingObjectIdentifier>\n  </premis:event>\n</premis:premis>'
        fo.write(string_log)

def configure_sip_premis(premis_path, id, tarfolder):
    print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Configuring sip premis.xml...')
    start_premis = f'<?xml version=\'1.0\' encoding=\'UTF-8\'?>\n <premis:premis xmlns:premis="http://arkivverket.no/standarder/PREMIS" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink" xsi:schemaLocation="http://arkivverket.no/standarder/PREMIS http://schema.arkivverket.no/PREMIS/v2.0/DIAS_PREMIS.xsd" version="2.0">\n  <premis:object xsi:type="premis:file">\n    <premis:objectIdentifier>\n      <premis:objectIdentifierType>NO/RA</premis:objectIdentifierType>\n      <premis:objectIdentifierValue>{id}</premis:objectIdentifierValue>\n    </premis:objectIdentifier>\n    <premis:preservationLevel>\n      <premis:preservationLevelValue>full</premis:preservationLevelValue>\n    </premis:preservationLevel>\n    <premis:objectCharacteristics>\n      <premis:compositionLevel>0</premis:compositionLevel>\n      <premis:format>\n        <premis:formatDesignation>\n          <premis:formatName>tar</premis:formatName>\n        </premis:formatDesignation>\n      </premis:format>\n    </premis:objectCharacteristics>\n    <premis:storage>\n      <premis:storageMedium>ESSArch Tools</premis:storageMedium>\n    </premis:storage>\n  </premis:object>\n'
    end_premis = f'  <premis:agent>\n    <premis:agentIdentifier>\n      <premis:agentIdentifierType>NO/RA</premis:agentIdentifierType>\n      <premis:agentIdentifierValue>ESSArch</premis:agentIdentifierValue>\n    </premis:agentIdentifier>\n    <premis:agentName>ESSArch Tools</premis:agentName>\n    <premis:agentType>software</premis:agentType>\n  </premis:agent>\n</premis:premis>'
    with open(premis_path, "w", encoding="utf-8") as fo:
        fo.write(start_premis)
        for root, _, files in os.walk(tarfolder):    
            for file in files:
                tmp_path = root.split("/")[4].replace("\\", "/") + "/" + file
                if tmp_path != f'{id}/mets.xml' and tmp_path != f'{id}/administrative_metadata/premis.xml':
                    if tmp_path not in INFO_DICT.keys():
                        sha = hashlib.sha256()
                        with open(os.path.join(root, file), "rb") as f:
                            tmp_data = f.read(4000000)
                            tmp_magic = magic.from_buffer(tmp_data, mime=True)
                            while len(tmp_data)>0:
                                sha.update(tmp_data)
                                tmp_data = f.read(4000000)
                        INFO_DICT[tmp_path] = [sha.hexdigest(), tmp_magic]
                    fill_premis = f'  <premis:object xsi:type="premis:file">\n    <premis:objectIdentifier>\n      <premis:objectIdentifierType>NO/RA</premis:objectIdentifierType>\n      <premis:objectIdentifierValue>{tmp_path}</premis:objectIdentifierValue>\n    </premis:objectIdentifier>\n    <premis:objectCharacteristics>\n      <premis:compositionLevel>0</premis:compositionLevel>\n      <premis:fixity>\n        <premis:messageDigestAlgorithm>SHA-256</premis:messageDigestAlgorithm>\n        <premis:messageDigest>{INFO_DICT[tmp_path][0]}</premis:messageDigest>\n        <premis:messageDigestOriginator>ESSArch</premis:messageDigestOriginator>\n      </premis:fixity>\n      <premis:size>{os.stat(os.path.join(root, file)).st_size}</premis:size>\n      <premis:format>\n        <premis:formatDesignation>\n          <premis:formatName>{os.path.splitext(file)[1][1:]}</premis:formatName>\n        </premis:formatDesignation>\n      </premis:format>\n    </premis:objectCharacteristics>\n    <premis:storage>\n      <premis:contentLocation>\n        <premis:contentLocationType>SIP</premis:contentLocationType>\n        <premis:contentLocationValue>{id}</premis:contentLocationValue>\n      </premis:contentLocation>\n    </premis:storage>\n    <premis:relationship>\n      <premis:relationshipType>structural</premis:relationshipType>\n      <premis:relationshipSubType>is part of</premis:relationshipSubType>\n      <premis:relatedObjectIdentification>\n        <premis:relatedObjectIdentifierType>NO/RA</premis:relatedObjectIdentifierType>\n        <premis:relatedObjectIdentifierValue>{id}</premis:relatedObjectIdentifierValue>\n      </premis:relatedObjectIdentification>\n    </premis:relationship>\n  </premis:object>\n'
                    fo.write(fill_premis)
        fo.write(end_premis)

def configure_sip_mets(mets_path, id, tarfolder, creation_date, premis_path): #TODO
    print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Configuring sip mets.xml...')
    with open(mets_path, "w", encoding="utf-8") as fo:
        sha = hashlib.sha256()
        with open(premis_path, "rb") as f:
            tmp_data = f.read(4000000)
            while len(tmp_data)>0:
                sha.update(tmp_data)
                tmp_data = f.read(4000000)
        id_list = [f'ID{uuid1()}']
        start_mets = f'<?xml version="1.0" encoding="UTF-8"?>\n<mets:mets xmlns:mets="http://www.loc.gov/METS/" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/METS/ http://schema.arkivverket.no/METS/mets.xsd" PROFILE="http://xml.ra.se/METS/RA_METS_eARD.xml" LABEL="{label_entry.get()}" TYPE="SIP" ID="ID{uuid1()}" OBJID="UUID:{id}">\n    <mets:metsHdr CREATEDATE="{creation_date}" RECORDSTATUS="NEW">\n        <mets:agent TYPE="ORGANIZATION" ROLE="ARCHIVIST">\n            <mets:name>{archivist_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{system_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{system_ver_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{type_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="CREATOR">\n            <mets:name>{creator_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_org_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="INDIVIDUAL" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_pers_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_software_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="OTHER" OTHERROLE="SUBMITTER">\n            <mets:name>{submitter_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="INDIVIDUAL" ROLE="OTHER" OTHERROLE="SUBMITTER">\n            <mets:name>{submitter_pers_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="IPOWNER">\n            <mets:name>{owner_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="PRESERVATION">\n            <mets:name>{preserver_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:altRecordID TYPE="SUBMISSIONAGREEMENT">{submission_entry.get()}</mets:altRecordID>\n        <mets:altRecordID TYPE="STARTDATE">{period_start_entry.get()}</mets:altRecordID>\n        <mets:altRecordID TYPE="ENDDATE">{period_end_entry.get()}</mets:altRecordID>\n        <mets:metsDocumentID>mets.xml</mets:metsDocumentID>\n    </mets:metsHdr>\n    <mets:amdSec ID="amdSec001">\n        <mets:digiprovMD ID="digiprovMD001">\n            <mets:mdRef MIMETYPE="text/xml" CHECKSUMTYPE="SHA-256" CHECKSUM="{sha.hexdigest()}" MDTYPE="PREMIS" xlink:href="file:administrative_metadata/premis.xml" LOCTYPE="URL" CREATED="{datetime.fromtimestamp(os.path.getmtime(premis_path)).strftime("%Y-%m-%dT%H:%M:%S+02:00")}" xlink:type="simple" ID="{id_list[-1]}" SIZE="{os.stat(premis_path).st_size}"/>\n        </mets:digiprovMD>\n    </mets:amdSec>\n    <mets:fileSec>\n        <mets:fileGrp ID="fgrp001" USE="FILES">\n'
        end_mets = f'            </mets:div>\n        </mets:div>\n    </mets:structMap>\n</mets:mets>'
        fo.write(start_mets)
        for root, _, files in os.walk(tarfolder):
            for file in files:
                tmp_path = "file:" + (root.replace("\\", "/") + "/" + file).removeprefix(f'{tarfolder}/')
                if tmp_path != "file:administrative_metadata/premis.xml" and tmp_path != "file:mets.xml":
                    id_list.append(f'ID{uuid1()}')
                    fill_mets = f'            <mets:file MIMETYPE="{INFO_DICT[id+"/"+tmp_path[5:]][1]}" CHECKSUMTYPE="SHA-256" CREATED="{datetime.fromtimestamp(os.path.getmtime(os.path.join(root,file))).strftime("%Y-%m-%dT%H:%M:%S+02:00")}" CHECKSUM="{INFO_DICT[id+"/"+tmp_path[5:]][0]}" USE="Datafile" ID="{id_list[-1]}" SIZE="{os.stat(os.path.join(root,file)).st_size}">\n                <mets:FLocat xlink:href="{tmp_path}" LOCTYPE="URL" xlink:type="simple"/>\n            </mets:file>\n'
                    fo.write(fill_mets)
        fill_mets = f'        </mets:fileGrp>\n    </mets:fileSec>\n    <mets:structMap>\n        <mets:div LABEL="Package">\n            <mets:div ADMID="amdSec001" LABEL="Content Description">\n                <mets:fptr FILEID="{id_list.pop(0)}"/>\n            </mets:div>\n            <mets:div ADMID="amdSec001" LABEL="Datafiles">\n'
        fo.write(fill_mets)
        while id_list:
            fill_mets = f'                <mets:fptr FILEID="{id_list.pop(0)}"/>\n'
            fo.write(fill_mets)
        fo.write(end_mets)

def configure_sip_info(info_path, tar_path, id, creation_date):
    print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Configuring sip info.xml...')
    extra_id = f'ID{uuid1()}'
    sha = hashlib.sha256()
    with open(tar_path, "rb") as f:
        tmp_data = f.read(4000000)
        while len(tmp_data)>0:
            sha.update(tmp_data)
            tmp_data = f.read(4000000)
    with open(info_path, "w", encoding="utf-8") as fo:
        string_info = f'<?xml version="1.0" encoding="UTF-8"?>\n<mets:mets xmlns:mets="http://www.loc.gov/METS/" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/METS/ http://schema.arkivverket.no/METS/info.xsd" PROFILE="http://xml.ra.se/METS/RA_METS_eARD.xml" LABEL="{label_entry.get()}" TYPE="SIP" ID="ID{uuid1()}" OBJID="UUID:{id}">\n    <mets:metsHdr CREATEDATE="{creation_date}" RECORDSTATUS="NEW">\n        <mets:agent TYPE="ORGANIZATION" ROLE="ARCHIVIST">\n            <mets:name>{archivist_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{system_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{system_ver_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{type_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="CREATOR">\n            <mets:name>{creator_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_org_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="INDIVIDUAL" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_pers_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_software_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="OTHER" OTHERROLE="SUBMITTER">\n            <mets:name>{submitter_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="INDIVIDUAL" ROLE="OTHER" OTHERROLE="SUBMITTER">\n            <mets:name>{submitter_pers_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="IPOWNER">\n            <mets:name>{owner_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="PRESERVATION">\n            <mets:name>{preserver_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:altRecordID TYPE="SUBMISSIONAGREEMENT">{submission_entry.get()}</mets:altRecordID>\n        <mets:altRecordID TYPE="STARTDATE">{period_start_entry.get()}</mets:altRecordID>\n        <mets:altRecordID TYPE="ENDDATE">{period_end_entry.get()}</mets:altRecordID>\n        <mets:metsDocumentID>info.xml</mets:metsDocumentID>\n    </mets:metsHdr>\n    <mets:fileSec>\n        <mets:fileGrp ID="fgrp001" USE="FILES">\n            <mets:file MIMETYPE="application/x-tar" CHECKSUMTYPE="SHA-256" CREATED="{datetime.fromtimestamp(os.path.getmtime(tar_path)).strftime("%Y-%m-%dT%H:%M:%S+02:00")}" CHECKSUM="{sha.hexdigest()}" USE="Datafile" ID="{extra_id}" SIZE="{os.stat(tar_path).st_size}">\n                <mets:FLocat xlink:href="file:{os.path.basename(tar_path)}" LOCTYPE="URL" xlink:type="simple"/>\n            </mets:file>\n        </mets:fileGrp>\n    </mets:fileSec>\n    <mets:structMap>\n        <mets:div LABEL="Package">\n            <mets:div LABEL="Content Description"/>\n            <mets:div LABEL="Datafiles">\n                <mets:fptr FILEID="{extra_id}"/>\n            </mets:div>\n        </mets:div>\n    </mets:structMap>\n</mets:mets>'
        fo.write(string_info)

def configure_aic_log(log_path, aic_id, sip_id, create_date):
    print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Configuring aic log.xml...')
    with open(log_path, "w", encoding="utf-8") as fo:
        string_log = f'<?xml version=\'1.0\' encoding=\'UTF-8\'?>\n<premis:premis xmlns:premis="http://arkivverket.no/standarder/PREMIS" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink" xsi:schemaLocation="http://arkivverket.no/standarder/PREMIS http://schema.arkivverket.no/PREMIS/v2.0/DIAS_PREMIS.xsd" version="2.0">\n  <premis:object xsi:type="premis:file">\n    <premis:objectIdentifier>\n      <premis:objectIdentifierType>NO/RA</premis:objectIdentifierType>\n      <premis:objectIdentifierValue>{sip_id}</premis:objectIdentifierValue>\n    </premis:objectIdentifier>\n    <premis:preservationLevel>\n      <premis:preservationLevelValue>full</premis:preservationLevelValue>\n    </premis:preservationLevel>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>aic_object</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{aic_id}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>createdate</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{create_date}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>archivist_organization</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{archivist_org_combo.get()}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>label</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{label_entry.get()}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>iptype</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>SIP</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:objectCharacteristics>\n      <premis:compositionLevel>0</premis:compositionLevel>\n      <premis:format>\n        <premis:formatDesignation>\n          <premis:formatName>tar</premis:formatName>\n        </premis:formatDesignation>\n      </premis:format>\n    </premis:objectCharacteristics>\n    <premis:storage>\n      <premis:storageMedium>Preservation platform ESSArch</premis:storageMedium>\n    </premis:storage>\n    <premis:relationship>\n      <premis:relationshipType>structural</premis:relationshipType>\n      <premis:relationshipSubType>is part of</premis:relationshipSubType>\n      <premis:relatedObjectIdentification>\n        <premis:relatedObjectIdentifierType>NO/RA</premis:relatedObjectIdentifierType>\n        <premis:relatedObjectIdentifierValue>{aic_id}</premis:relatedObjectIdentifierValue>\n      </premis:relatedObjectIdentification>\n    </premis:relationship>\n  </premis:object>\n  <premis:event>\n    <premis:eventIdentifier>\n      <premis:eventIdentifierType>NO/RA</premis:eventIdentifierType>\n      <premis:eventIdentifierValue>{uuid1()}</premis:eventIdentifierValue>\n    </premis:eventIdentifier>\n    <premis:eventType>20000</premis:eventType>\n    <premis:eventDateTime>{create_date}</premis:eventDateTime>\n    <premis:eventDetail>Created log circular</premis:eventDetail>\n    <premis:eventOutcomeInformation>\n      <premis:eventOutcome>0</premis:eventOutcome>\n      <premis:eventOutcomeDetail>\n        <premis:eventOutcomeDetailNote>Success to create logfile</premis:eventOutcomeDetailNote>\n      </premis:eventOutcomeDetail>\n    </premis:eventOutcomeInformation>\n    <premis:linkingAgentIdentifier>\n      <premis:linkingAgentIdentifierType>NO/RA</premis:linkingAgentIdentifierType>\n      <premis:linkingAgentIdentifierValue>admin</premis:linkingAgentIdentifierValue>\n    </premis:linkingAgentIdentifier>\n    <premis:linkingObjectIdentifier>\n      <premis:linkingObjectIdentifierType>NO/RA</premis:linkingObjectIdentifierType>\n      <premis:linkingObjectIdentifierValue>{sip_id}</premis:linkingObjectIdentifierValue>\n    </premis:linkingObjectIdentifier>\n  </premis:event>\n</premis:premis>'
        fo.write(string_log)

#Main function
def main_func():
    if content_path_label.cget("text") and all(len(i.get()) != 0 for i in TEXT_LIST):
        print(f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: Building structure...')
        sip_id = uuid1()

        output_folder = 1
        while(os.path.isdir(f'./{output_folder}')):
            output_folder += 1
        output_folder = f'./{output_folder}'
        tarfile = f'{output_folder}/{sip_id}/content/{sip_id}'

        os.makedirs(f'{output_folder}/{sip_id}/administrative_metadata/repository_operations')
        os.makedirs(f'{output_folder}/{sip_id}/descriptive_metadata')
        os.makedirs(f'{tarfile}/administrative_metadata')
        os.makedirs(f'{tarfile}/descriptive_metadata')
        os.makedirs(f'{tarfile}/content')

        shutil.copy(os.path.join(sys._MEIPASS, "files/mets.xsd"), f'{tarfile}/mets.xsd')
        shutil.copy(os.path.join(sys._MEIPASS, "files/DIAS_PREMIS.xsd"), f'{tarfile}/administrative_metadata/DIAS_PREMIS.xsd')
        shutil.copytree(os.path.abspath(content_path_label.cget("text")), f'{tarfile}/content', copy_function=shutil.copy, dirs_exist_ok=True)
        if descriptive_path_label.cget("text"):
            shutil.copytree(os.path.abspath(descriptive_path_label.cget("text")), f'{tarfile}/descriptive_metadata', copy_function=shutil.copy, dirs_exist_ok=True)
        if administrative_path_label.cget("text"):
            shutil.copytree(os.path.abspath(administrative_path_label.cget("text")), f'{tarfile}/administrative_metadata', copy_function=shutil.copy, dirs_exist_ok=True)

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
    else:
        messagebox.showerror("Error", "Non-optional input missing.")

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
window.geometry('{width}x{height}+{pos_right}+{pos_down}'.format(width=(window.winfo_screenwidth() // 2)+(window.winfo_screenwidth() // 5), height=(window.winfo_screenheight() // 2)+(window.winfo_screenheight() // 5), pos_right=(window.winfo_screenwidth() // 4)-(window.winfo_screenwidth() // 10), pos_down=(window.winfo_screenheight() // 4)-(window.winfo_screenheight() // 10)))
window.config(background="teal")

#Configures columns and rows in the main window's grid
for i in range(21):
    window.columnconfigure(i, weight=1, pad=0)
    window.rowconfigure(i, weight=1, pad=0)

#Creates labels, combo-boxes, entrys and some buttons with self-explanatory names
TEXT_LIST = [StringVar(window, name=f'{i}') for i in range(16)]
content_path_label = Label(window, text="", relief="solid")
content_path_button = Button(window, text="Browse Content", command=lambda: browse_files(content_path_label))
descriptive_path_label = Label(window, text="", relief="solid")
descriptive_path_button = Button(window, text="Browse Descriptive Metadata (optional)", command=lambda: browse_files(descriptive_path_label))
administrative_path_label = Label(window, text="", relief="solid")
administrative_path_button = Button(window, text="Browse Administrative Metadata (optional)", command=lambda: browse_files(administrative_path_label))
import_button = Button(window, text="Import mets.xml Metadata (optional)", command=lambda: import_metadata(filedialog.askopenfile(initialdir="./", title="Choose metadata file", filetypes=[("XML files", "*.xml")])))
system_label = Label(window, text="System:", background="black", foreground="white")
system_combo = ttk.Combobox(window, textvariable=TEXT_LIST[0], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"), postcommand=lambda: combo_helper(system_combo,SYSTEM_LIST))
system_ver_label = Label(window, text="System Version:", background="black", foreground="white")
system_ver_entry = Entry(window, textvariable=TEXT_LIST[1], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
submission_label = Label(window, text="Submission Agreement:", background="black", foreground="white")
submission_entry = Entry(window, textvariable=TEXT_LIST[2], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
archivist_org_label = Label(window, text="Archivist Organization:", background="black", foreground="white")
archivist_org_combo = ttk.Combobox(window, textvariable=TEXT_LIST[3], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"), postcommand=lambda: combo_helper(archivist_org_combo,MUNICIPALITY_LIST))
label_label = Label(window, text="Label:", background="black", foreground="white")
label_entry = Entry(window, textvariable=TEXT_LIST[4], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
type_label = Label(window, text="Archivist System Type:", background="black", foreground="white")
type_combo = ttk.Combobox(window, textvariable=TEXT_LIST[5], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"), postcommand=lambda: combo_helper(type_combo,["SIARD", "NOARK-5", "Postjournaler"]))
owner_org_label = Label(window, text="Owner Organization:", background="black", foreground="white")
owner_org_combo = ttk.Combobox(window, textvariable=TEXT_LIST[6], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"), postcommand=lambda: combo_helper(owner_org_combo,MUNICIPALITY_LIST))
producer_org_label = Label(window, text="Producer Organization:", background="black", foreground="white")
producer_org_entry = Entry(window, textvariable=TEXT_LIST[7], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
producer_pers_label = Label(window, text="Producer Person:", background="black", foreground="white")
producer_pers_entry = Entry(window, textvariable=TEXT_LIST[8], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
producer_software_label = Label(window, text="Producer Software:", background="black", foreground="white")
producer_software_entry = Entry(window, textvariable=TEXT_LIST[9], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
period_start_label = Label(window, text="Period Start:", background="black", foreground="white")
period_start_entry = Entry(window, textvariable=TEXT_LIST[10], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
period_end_label = Label(window, text="Period End:", background="black", foreground="white")
period_end_entry = Entry(window, textvariable=TEXT_LIST[11], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
submitter_org_label = Label(window, text="Submitter Organization:", background="black", foreground="white")
submitter_org_combo = ttk.Combobox(window, textvariable=TEXT_LIST[12], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"), postcommand=lambda: combo_helper(submitter_org_combo,MUNICIPALITY_LIST))
submitter_pers_label = Label(window, text="Submitter Person:", background="black", foreground="white")
submitter_pers_entry = Entry(window, textvariable=TEXT_LIST[13], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
creator_label = Label(window, text="Creator:", background="black", foreground="white")
creator_entry = Entry(window, textvariable=TEXT_LIST[14], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
preserver_label = Label(window, text="Preserver:", background="black", foreground="white")
preserver_entry = Entry(window, textvariable=TEXT_LIST[15], validate='key', validatecommand=(window.register(validation_func), '%S', "%d"))
button_execute = Button(window, text="Create Dias Package", command=main_func)

#Adds the widgets to the main window
content_path_label.grid(column=1, row=1, columnspan=13, sticky="NSEW")
content_path_button.grid(column=16, row=1, columnspan=4, sticky="NSEW")
descriptive_path_label.grid(column=1, row=2, columnspan=13, sticky="NSEW")
descriptive_path_button.grid(column=16, row=2, columnspan=4, sticky="NSEW")
administrative_path_label.grid(column=1, row=3, columnspan=13, sticky="NSEW")
administrative_path_button.grid(column=16, row=3, columnspan=4, sticky="NSEW")
import_button.grid(column=1, row=4, columnspan=19, sticky="NSEW")
system_label.grid(column=1, row=6, columnspan=2, sticky="NSEW")
system_combo.grid(column=3, row=6, columnspan=7, sticky="NSEW")
system_ver_label.grid(column=10, row=6, columnspan=2, sticky="NSEW")
system_ver_entry.grid(column=12, row=6, columnspan=8, sticky="NSEW")
submission_label.grid(column=1, row=7, columnspan=2, sticky="NSEW")
submission_entry.grid(column=3, row=7, columnspan=17, sticky="NSEW")
archivist_org_label.grid(column=1, row=8, columnspan=2, sticky="NSEW")
archivist_org_combo.grid(column=3, row=8, columnspan=17, sticky="NSEW")
label_label.grid(column=1, row=9, columnspan=2, sticky="NSEW")
label_entry.grid(column=3, row=9, columnspan=17, sticky="NSEW")
type_label.grid(column=1, row=10, columnspan=2, sticky="NSEW")
type_combo.grid(column=3, row=10, columnspan=17, sticky="NSEW")
owner_org_label.grid(column=1, row=11, columnspan=2, sticky="NSEW")
owner_org_combo.grid(column=3, row=11, columnspan=17, sticky="NSEW")
producer_org_label.grid(column=1, row=12, columnspan=2, sticky="NSEW")
producer_org_entry.grid(column=3, row=12, columnspan=17, sticky="NSEW")
producer_pers_label.grid(column=1, row=13, columnspan=2, sticky="NSEW")
producer_pers_entry.grid(column=3, row=13, columnspan=17, sticky="NSEW")
producer_software_label.grid(column=1, row=14, columnspan=2, sticky="NSEW")
producer_software_entry.grid(column=3, row=14, columnspan=17, sticky="NSEW")
period_start_label.grid(column=1, row=15, columnspan=2, sticky="NSEW")
period_start_entry.grid(column=3, row=15, columnspan=7, sticky="NSEW")
period_end_label.grid(column=10, row=15, columnspan=2, sticky="NSEW")
period_end_entry.grid(column=12, row=15, columnspan=8, sticky="NSEW")
submitter_org_label.grid(column=1, row=16, columnspan=2, sticky="NSEW")
submitter_org_combo.grid(column=3, row=16, columnspan=7, sticky="NSEW")
submitter_pers_label.grid(column=10, row=16, columnspan=2, sticky="NSEW")
submitter_pers_entry.grid(column=12, row=16, columnspan=8, sticky="NSEW")
creator_label.grid(column=1, row=17, columnspan=2, sticky="NSEW")
creator_entry.grid(column=3, row=17, columnspan=7, sticky="NSEW")
preserver_label.grid(column=10, row=17, columnspan=2, sticky="NSEW")
preserver_entry.grid(column=12, row=17, columnspan=8, sticky="NSEW")
button_execute.grid(column=1, row=19, columnspan=19, sticky="NSEW")

#Runs the main window and keeps it going
window.mainloop()
