import os
import sys
import magic
import shutil
import hashlib
import tkinter
import threading
import traceback
import customtkinter
import xml.etree.ElementTree as ET
from uuid import uuid1
from datetime import datetime
from subprocess import run, DEVNULL
from tkinter import messagebox, filedialog, Label, StringVar, Menu

MUNICIPALITY_LIST = sorted(["5041 Snåsa Kommune", "5057 Ørland Kommune", "5059 Orkland Kommune", "5034 Meråker Kommune", "5037 Levanger Kommune", "5025 Røros Kommune", "5016 Agdenes Kommune", "5012 Snillfjord Kommune", "5036 Frosta Kommune", "5023 Meldal Kommune", "5044 Namsskogan Kommune", "5043 Røyrvik Kommune", "5011 Hemne Kommune", "5032 Selbu Kommune", "5035 Stjørdal Kommune", "5046 Høylandet Kommune", "5042 Lierne Kommune", "5045 Grong Kommune","5049 Flatanger Kommune","5014 Frøya Kommune","5055 Heim Kommune","5013 Hitra Kommune","5026 Holtålen Kommune","5053 Inderøy Kommune","5054 Indre Fosen Kommune","5031 Malvik Kommune","5028 Melhus Kommune","5027 Midtre Gauldal Kommune","5005 Namsos Kommune","5060 Nærøysund Kommune","5021 Oppdal Kommune","3430 Os Kommune","5047 Overhalla Kommune","5020 Osen Kommune","5022 Rennebu Kommune","5029 Skaun Kommune","5006 Steinkjer Kommune","5033 Tydal Kommune","5038 Verdal Kommune","5058 Åfjord Kommune"], key=lambda x: x.split(" ")[1])
SYSTEM_LIST = sorted(["ESA", "Visma Velferd", "Visma Familia", "Visma HsPro", "WinMed Helse", "Ephorte", "Visma Flyt Skole", "Visma Profil", "SystemX", "P360", "Digora", "Oppad", "CGM Helsestasjon", "Visma Flyt Sampro", "Gerica", "Socio"])
USERNAME = "admin"

#Function to browse computer for files
def browse_files(label: Label):
    file = filedialog.askdirectory(initialdir="./", title="Choose a folder whose content should be packaged")
    label.configure(text=file)

#Function to import metadata from chosen mets.xml file
def import_metadata(path: str):
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
                            field = compare_dict[frozenset(i.attrib.values())].pop(0)
                            field.set(i[j].text)
                elif i.text and frozenset(i.attrib.values()) in compare_dict:
                    compare_dict[frozenset(i.attrib.values())][0].set(i.text)
        else:
            messagebox.showerror("Error", "Unable to recognise file as a mets.xml file.")

#Gets the sha-256 hash, mimetype, filesize and file creation date of all files in target directory
def gather_file_info(directory: str, prefix: str) -> dict:
    logging("Gathering checksums...")
    info_dict = {}
    for root, _, files in os.walk(directory):
        for file in files:
            sha = hashlib.sha256()
            with open(os.path.join(root, file), "rb") as f:
                tmp_data = f.read(4000000)
                tmp_magic = magic.from_buffer(tmp_data, mime=True)
                while tmp_data:
                    sha.update(tmp_data)
                    tmp_data = f.read(4000000)
            info_dict[f'{prefix}{root.removeprefix(directory)}/{file}'.replace("\\", "/")] = [sha.hexdigest(), tmp_magic, os.stat(os.path.join(root, file)).st_size, datetime.fromtimestamp(os.path.getmtime(os.path.join(root,file))).strftime("%Y-%m-%dT%H:%M:%S+02:00")]
    return info_dict

#Packages the sip into a tarfile
def pack_sip(sip_tarfile: str, id: str, content_path: str):
    logging("Packaging sip...")
    run(f'{os.environ["PROGRAMFILES"]}\\7-Zip\\7z.exe a {sip_tarfile}.tar {sip_tarfile} -aou -sdel', stdout=DEVNULL, stderr=DEVNULL)
    run(f'{os.environ["PROGRAMFILES"]}\\7-Zip\\7z.exe a {sip_tarfile}.tar "{content_path}"', stdout=DEVNULL, stderr=DEVNULL)
    run(f'{os.environ["PROGRAMFILES"]}\\7-Zip\\7z.exe rn {sip_tarfile}.tar -r "{os.path.basename(content_path)}\\" {id}\\content\\', stdout=DEVNULL, stderr=DEVNULL)

def configure_sip_log(log_path: str, id: str, create_date: str):
    logging("Configuring sip log.xml...")
    with open(log_path, "w", encoding="utf-8") as fo:
        string_log = f'<?xml version=\'1.0\' encoding=\'UTF-8\'?>\n<premis:premis xmlns:premis="http://arkivverket.no/standarder/PREMIS" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink" xsi:schemaLocation="http://arkivverket.no/standarder/PREMIS http://schema.arkivverket.no/PREMIS/v2.0/DIAS_PREMIS.xsd" version="2.0">\n  <premis:object xsi:type="premis:file">\n    <premis:objectIdentifier>\n      <premis:objectIdentifierType>NO/RA</premis:objectIdentifierType>\n      <premis:objectIdentifierValue>{id}</premis:objectIdentifierValue>\n    </premis:objectIdentifier>\n    <premis:preservationLevel>\n      <premis:preservationLevelValue>full</premis:preservationLevelValue>\n    </premis:preservationLevel>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>aic_object</premis:significantPropertiesType>\n      <premis:significantPropertiesValue></premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>createdate</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{create_date}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>archivist_organization</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{archivist_org_combo.get()}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>label</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{label_entry.get()}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>iptype</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>SIP</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:objectCharacteristics>\n      <premis:compositionLevel>0</premis:compositionLevel>\n      <premis:format>\n        <premis:formatDesignation>\n          <premis:formatName>tar</premis:formatName>\n        </premis:formatDesignation>\n      </premis:format>\n    </premis:objectCharacteristics>\n    <premis:storage>\n      <premis:storageMedium>Preservation platform ESSArch</premis:storageMedium>\n    </premis:storage>\n    <premis:relationship>\n      <premis:relationshipType>structural</premis:relationshipType>\n      <premis:relationshipSubType>is part of</premis:relationshipSubType>\n      <premis:relatedObjectIdentification>\n        <premis:relatedObjectIdentifierType>NO/RA</premis:relatedObjectIdentifierType>\n        <premis:relatedObjectIdentifierValue></premis:relatedObjectIdentifierValue>\n      </premis:relatedObjectIdentification>\n    </premis:relationship>\n  </premis:object>\n  <premis:event>\n    <premis:eventIdentifier>\n      <premis:eventIdentifierType>NO/RA</premis:eventIdentifierType>\n      <premis:eventIdentifierValue>{uuid1()}</premis:eventIdentifierValue>\n    </premis:eventIdentifier>\n    <premis:eventType>10000</premis:eventType>\n    <premis:eventDateTime>{create_date}</premis:eventDateTime>\n    <premis:eventDetail>Log circular created</premis:eventDetail>\n    <premis:eventOutcomeInformation>\n      <premis:eventOutcome>0</premis:eventOutcome>\n      <premis:eventOutcomeDetail>\n        <premis:eventOutcomeDetailNote>Success to create logfile</premis:eventOutcomeDetailNote>\n      </premis:eventOutcomeDetail>\n    </premis:eventOutcomeInformation>\n    <premis:linkingAgentIdentifier>\n      <premis:linkingAgentIdentifierType>NO/RA</premis:linkingAgentIdentifierType>\n      <premis:linkingAgentIdentifierValue>{USERNAME}</premis:linkingAgentIdentifierValue>\n    </premis:linkingAgentIdentifier>\n    <premis:linkingObjectIdentifier>\n      <premis:linkingObjectIdentifierType>NO/RA</premis:linkingObjectIdentifierType>\n      <premis:linkingObjectIdentifierValue>{id}</premis:linkingObjectIdentifierValue>\n    </premis:linkingObjectIdentifier>\n  </premis:event>\n</premis:premis>'
        fo.write(string_log)

def configure_sip_premis(premis_path: str, id: str, info_dict: dict):
    logging("Configuring sip premis.xml...")
    start_premis = f'<?xml version=\'1.0\' encoding=\'UTF-8\'?>\n <premis:premis xmlns:premis="http://arkivverket.no/standarder/PREMIS" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink" xsi:schemaLocation="http://arkivverket.no/standarder/PREMIS http://schema.arkivverket.no/PREMIS/v2.0/DIAS_PREMIS.xsd" version="2.0">\n  <premis:object xsi:type="premis:file">\n    <premis:objectIdentifier>\n      <premis:objectIdentifierType>NO/RA</premis:objectIdentifierType>\n      <premis:objectIdentifierValue>{id}</premis:objectIdentifierValue>\n    </premis:objectIdentifier>\n    <premis:preservationLevel>\n      <premis:preservationLevelValue>full</premis:preservationLevelValue>\n    </premis:preservationLevel>\n    <premis:objectCharacteristics>\n      <premis:compositionLevel>0</premis:compositionLevel>\n      <premis:format>\n        <premis:formatDesignation>\n          <premis:formatName>tar</premis:formatName>\n        </premis:formatDesignation>\n      </premis:format>\n    </premis:objectCharacteristics>\n    <premis:storage>\n      <premis:storageMedium>ESSArch Tools</premis:storageMedium>\n    </premis:storage>\n  </premis:object>\n'
    end_premis = f'  <premis:agent>\n    <premis:agentIdentifier>\n      <premis:agentIdentifierType>NO/RA</premis:agentIdentifierType>\n      <premis:agentIdentifierValue>ESSArch</premis:agentIdentifierValue>\n    </premis:agentIdentifier>\n    <premis:agentName>ESSArch Tools</premis:agentName>\n    <premis:agentType>software</premis:agentType>\n  </premis:agent>\n</premis:premis>'
    with open(premis_path, "w", encoding="utf-8") as fo:
        fo.write(start_premis)
        for path, info in info_dict.items():
            if path != f'{id}/mets.xml' and path != f'{id}/administrative_metadata/premis.xml':
                fill_premis = f'  <premis:object xsi:type="premis:file">\n    <premis:objectIdentifier>\n      <premis:objectIdentifierType>NO/RA</premis:objectIdentifierType>\n      <premis:objectIdentifierValue>{path}</premis:objectIdentifierValue>\n    </premis:objectIdentifier>\n    <premis:objectCharacteristics>\n      <premis:compositionLevel>0</premis:compositionLevel>\n      <premis:fixity>\n        <premis:messageDigestAlgorithm>SHA-256</premis:messageDigestAlgorithm>\n        <premis:messageDigest>{info[0]}</premis:messageDigest>\n        <premis:messageDigestOriginator>ESSArch</premis:messageDigestOriginator>\n      </premis:fixity>\n      <premis:size>{info[2]}</premis:size>\n      <premis:format>\n        <premis:formatDesignation>\n          <premis:formatName>{os.path.splitext(path)[1][1:]}</premis:formatName>\n        </premis:formatDesignation>\n      </premis:format>\n    </premis:objectCharacteristics>\n    <premis:storage>\n      <premis:contentLocation>\n        <premis:contentLocationType>SIP</premis:contentLocationType>\n        <premis:contentLocationValue>{id}</premis:contentLocationValue>\n      </premis:contentLocation>\n    </premis:storage>\n    <premis:relationship>\n      <premis:relationshipType>structural</premis:relationshipType>\n      <premis:relationshipSubType>is part of</premis:relationshipSubType>\n      <premis:relatedObjectIdentification>\n        <premis:relatedObjectIdentifierType>NO/RA</premis:relatedObjectIdentifierType>\n        <premis:relatedObjectIdentifierValue>{id}</premis:relatedObjectIdentifierValue>\n      </premis:relatedObjectIdentification>\n    </premis:relationship>\n  </premis:object>\n'
                fo.write(fill_premis)
        fo.write(end_premis)

def configure_sip_mets(mets_path: str, id: str, creation_date: str, premis_path: str, info_dict: dict):
    logging("Configuring sip mets.xml...")
    with open(mets_path, "w", encoding="utf-8") as fo:
        sha = hashlib.sha256()
        with open(premis_path, "rb") as f:
            while True:
                tmp_data = f.read(4000000)
                if not tmp_data:
                    break
                sha.update(tmp_data)
        id_list = [f'ID{uuid1()}']
        start_mets = f'<?xml version="1.0" encoding="UTF-8"?>\n<mets:mets xmlns:mets="http://www.loc.gov/METS/" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/METS/ http://schema.arkivverket.no/METS/mets.xsd" PROFILE="http://xml.ra.se/METS/RA_METS_eARD.xml" LABEL="{label_entry.get()}" TYPE="SIP" ID="ID{uuid1()}" OBJID="UUID:{id}">\n    <mets:metsHdr CREATEDATE="{creation_date}" RECORDSTATUS="NEW">\n        <mets:agent TYPE="ORGANIZATION" ROLE="ARCHIVIST">\n            <mets:name>{archivist_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{system_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{system_ver_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{type_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="CREATOR">\n            <mets:name>{creator_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_org_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="INDIVIDUAL" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_pers_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_software_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="OTHER" OTHERROLE="SUBMITTER">\n            <mets:name>{submitter_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="INDIVIDUAL" ROLE="OTHER" OTHERROLE="SUBMITTER">\n            <mets:name>{submitter_pers_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="IPOWNER">\n            <mets:name>{owner_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="PRESERVATION">\n            <mets:name>{preserver_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:altRecordID TYPE="SUBMISSIONAGREEMENT">{submission_entry.get()}</mets:altRecordID>\n        <mets:altRecordID TYPE="STARTDATE">{period_start_entry.get()}</mets:altRecordID>\n        <mets:altRecordID TYPE="ENDDATE">{period_end_entry.get()}</mets:altRecordID>\n        <mets:metsDocumentID>mets.xml</mets:metsDocumentID>\n    </mets:metsHdr>\n    <mets:amdSec ID="amdSec001">\n        <mets:digiprovMD ID="digiprovMD001">\n            <mets:mdRef MIMETYPE="text/xml" CHECKSUMTYPE="SHA-256" CHECKSUM="{sha.hexdigest()}" MDTYPE="PREMIS" xlink:href="file:administrative_metadata/premis.xml" LOCTYPE="URL" CREATED="{datetime.fromtimestamp(os.path.getmtime(premis_path)).strftime("%Y-%m-%dT%H:%M:%S+02:00")}" xlink:type="simple" ID="{id_list[-1]}" SIZE="{os.stat(premis_path).st_size}"/>\n        </mets:digiprovMD>\n    </mets:amdSec>\n    <mets:fileSec>\n        <mets:fileGrp ID="fgrp001" USE="FILES">\n'
        end_mets = f'            </mets:div>\n        </mets:div>\n    </mets:structMap>\n</mets:mets>'
        fo.write(start_mets)
        for path, info in info_dict.items():
            tmp_path = 'file:' + path.removeprefix(f'{id}/')
            if tmp_path != "file:administrative_metadata/premis.xml" and tmp_path != "file:mets.xml":
                id_list.append(f'ID{uuid1()}')
                fill_mets = f'            <mets:file MIMETYPE="{info[1]}" CHECKSUMTYPE="SHA-256" CREATED="{info[3]}" CHECKSUM="{info[0]}" USE="Datafile" ID="{id_list[-1]}" SIZE="{info[2]}">\n                <mets:FLocat xlink:href="{tmp_path}" LOCTYPE="URL" xlink:type="simple"/>\n            </mets:file>\n'
                fo.write(fill_mets)
        fill_mets = f'        </mets:fileGrp>\n    </mets:fileSec>\n    <mets:structMap>\n        <mets:div LABEL="Package">\n            <mets:div ADMID="amdSec001" LABEL="Content Description">\n                <mets:fptr FILEID="{id_list.pop(0)}"/>\n            </mets:div>\n            <mets:div ADMID="amdSec001" LABEL="Datafiles">\n'
        fo.write(fill_mets)
        while id_list:
            fill_mets = f'                <mets:fptr FILEID="{id_list.pop(0)}"/>\n'
            fo.write(fill_mets)
        fo.write(end_mets)

def configure_sip_info(info_path: str, tar_path: str, id: str, creation_date: str):
    logging("Configuring sip info.xml...")
    extra_id = f'ID{uuid1()}'
    sha = hashlib.sha256()
    with open(tar_path, "rb") as f:
        while True:
            tmp_data = f.read(4000000)
            if not tmp_data:
                break
            sha.update(tmp_data)
    with open(info_path, "w", encoding="utf-8") as fo:
        string_info = f'<?xml version="1.0" encoding="UTF-8"?>\n<mets:mets xmlns:mets="http://www.loc.gov/METS/" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/METS/ http://schema.arkivverket.no/METS/info.xsd" PROFILE="http://xml.ra.se/METS/RA_METS_eARD.xml" LABEL="{label_entry.get()}" TYPE="SIP" ID="ID{uuid1()}" OBJID="UUID:{id}">\n    <mets:metsHdr CREATEDATE="{creation_date}" RECORDSTATUS="NEW">\n        <mets:agent TYPE="ORGANIZATION" ROLE="ARCHIVIST">\n            <mets:name>{archivist_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{system_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{system_ver_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="ARCHIVIST">\n            <mets:name>{type_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="CREATOR">\n            <mets:name>{creator_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_org_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="INDIVIDUAL" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_pers_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="OTHER" OTHERTYPE="SOFTWARE" ROLE="OTHER" OTHERROLE="PRODUCER">\n            <mets:name>{producer_software_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="OTHER" OTHERROLE="SUBMITTER">\n            <mets:name>{submitter_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="INDIVIDUAL" ROLE="OTHER" OTHERROLE="SUBMITTER">\n            <mets:name>{submitter_pers_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="IPOWNER">\n            <mets:name>{owner_org_combo.get()}</mets:name>\n        </mets:agent>\n        <mets:agent TYPE="ORGANIZATION" ROLE="PRESERVATION">\n            <mets:name>{preserver_entry.get()}</mets:name>\n        </mets:agent>\n        <mets:altRecordID TYPE="SUBMISSIONAGREEMENT">{submission_entry.get()}</mets:altRecordID>\n        <mets:altRecordID TYPE="STARTDATE">{period_start_entry.get()}</mets:altRecordID>\n        <mets:altRecordID TYPE="ENDDATE">{period_end_entry.get()}</mets:altRecordID>\n        <mets:metsDocumentID>info.xml</mets:metsDocumentID>\n    </mets:metsHdr>\n    <mets:fileSec>\n        <mets:fileGrp ID="fgrp001" USE="FILES">\n            <mets:file MIMETYPE="application/x-tar" CHECKSUMTYPE="SHA-256" CREATED="{datetime.fromtimestamp(os.path.getmtime(tar_path)).strftime("%Y-%m-%dT%H:%M:%S+02:00")}" CHECKSUM="{sha.hexdigest()}" USE="Datafile" ID="{extra_id}" SIZE="{os.stat(tar_path).st_size}">\n                <mets:FLocat xlink:href="file:{os.path.basename(tar_path)}" LOCTYPE="URL" xlink:type="simple"/>\n            </mets:file>\n        </mets:fileGrp>\n    </mets:fileSec>\n    <mets:structMap>\n        <mets:div LABEL="Package">\n            <mets:div LABEL="Content Description"/>\n            <mets:div LABEL="Datafiles">\n                <mets:fptr FILEID="{extra_id}"/>\n            </mets:div>\n        </mets:div>\n    </mets:structMap>\n</mets:mets>'
        fo.write(string_info)

def configure_aic_log(log_path: str, aic_id: str, sip_id: str, create_date: str):
    logging("Configuring aic log.xml...")
    with open(log_path, "w", encoding="utf-8") as fo:
        string_log = f'<?xml version=\'1.0\' encoding=\'UTF-8\'?>\n<premis:premis xmlns:premis="http://arkivverket.no/standarder/PREMIS" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xlink="http://www.w3.org/1999/xlink" xsi:schemaLocation="http://arkivverket.no/standarder/PREMIS http://schema.arkivverket.no/PREMIS/v2.0/DIAS_PREMIS.xsd" version="2.0">\n  <premis:object xsi:type="premis:file">\n    <premis:objectIdentifier>\n      <premis:objectIdentifierType>NO/RA</premis:objectIdentifierType>\n      <premis:objectIdentifierValue>{sip_id}</premis:objectIdentifierValue>\n    </premis:objectIdentifier>\n    <premis:preservationLevel>\n      <premis:preservationLevelValue>full</premis:preservationLevelValue>\n    </premis:preservationLevel>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>aic_object</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{aic_id}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>createdate</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{create_date}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>archivist_organization</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{archivist_org_combo.get()}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>label</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>{label_entry.get()}</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:significantProperties>\n      <premis:significantPropertiesType>iptype</premis:significantPropertiesType>\n      <premis:significantPropertiesValue>SIP</premis:significantPropertiesValue>\n    </premis:significantProperties>\n    <premis:objectCharacteristics>\n      <premis:compositionLevel>0</premis:compositionLevel>\n      <premis:format>\n        <premis:formatDesignation>\n          <premis:formatName>tar</premis:formatName>\n        </premis:formatDesignation>\n      </premis:format>\n    </premis:objectCharacteristics>\n    <premis:storage>\n      <premis:storageMedium>Preservation platform ESSArch</premis:storageMedium>\n    </premis:storage>\n    <premis:relationship>\n      <premis:relationshipType>structural</premis:relationshipType>\n      <premis:relationshipSubType>is part of</premis:relationshipSubType>\n      <premis:relatedObjectIdentification>\n        <premis:relatedObjectIdentifierType>NO/RA</premis:relatedObjectIdentifierType>\n        <premis:relatedObjectIdentifierValue>{aic_id}</premis:relatedObjectIdentifierValue>\n      </premis:relatedObjectIdentification>\n    </premis:relationship>\n  </premis:object>\n  <premis:event>\n    <premis:eventIdentifier>\n      <premis:eventIdentifierType>NO/RA</premis:eventIdentifierType>\n      <premis:eventIdentifierValue>{uuid1()}</premis:eventIdentifierValue>\n    </premis:eventIdentifier>\n    <premis:eventType>20000</premis:eventType>\n    <premis:eventDateTime>{create_date}</premis:eventDateTime>\n    <premis:eventDetail>Created log circular</premis:eventDetail>\n    <premis:eventOutcomeInformation>\n      <premis:eventOutcome>0</premis:eventOutcome>\n      <premis:eventOutcomeDetail>\n        <premis:eventOutcomeDetailNote>Success to create logfile</premis:eventOutcomeDetailNote>\n      </premis:eventOutcomeDetail>\n    </premis:eventOutcomeInformation>\n    <premis:linkingAgentIdentifier>\n      <premis:linkingAgentIdentifierType>NO/RA</premis:linkingAgentIdentifierType>\n      <premis:linkingAgentIdentifierValue>{USERNAME}</premis:linkingAgentIdentifierValue>\n    </premis:linkingAgentIdentifier>\n    <premis:linkingObjectIdentifier>\n      <premis:linkingObjectIdentifierType>NO/RA</premis:linkingObjectIdentifierType>\n      <premis:linkingObjectIdentifierValue>{sip_id}</premis:linkingObjectIdentifierValue>\n    </premis:linkingObjectIdentifier>\n  </premis:event>\n</premis:premis>'
        fo.write(string_log)

#Main function
def main_func():
    tabview.set(3)
    PROGRESS_BAR.start()
    logging("Building structure...")
    try:
        sip_id = uuid1()
        output_folder = 1
        while(os.path.isdir(f'./{output_folder}')):
            output_folder += 1
        output_folder = f'./{output_folder}'
        tarfile = f'{output_folder}/{sip_id}/content/{sip_id}'

        #Build EPP-ready structure
        os.makedirs(f'{output_folder}/{sip_id}/administrative_metadata/repository_operations')
        os.makedirs(f'{output_folder}/{sip_id}/descriptive_metadata')
        os.makedirs(f'{tarfile}/administrative_metadata')
        os.makedirs(f'{tarfile}/descriptive_metadata')
        os.makedirs(f'{tarfile}/content')
        shutil.copy(os.path.join(sys._MEIPASS, "files/mets.xsd"), f'{tarfile}/mets.xsd')
        shutil.copy(os.path.join(sys._MEIPASS, "files/DIAS_PREMIS.xsd"), f'{tarfile}/administrative_metadata/DIAS_PREMIS.xsd')
        if descriptive_path_label.cget("text"):
            shutil.copytree(os.path.abspath(descriptive_path_label.cget("text")), f'{tarfile}/descriptive_metadata', copy_function=shutil.copy, dirs_exist_ok=True)
        if administrative_path_label.cget("text"):
            shutil.copytree(os.path.abspath(administrative_path_label.cget("text")), f'{tarfile}/administrative_metadata', copy_function=shutil.copy, dirs_exist_ok=True)

        #Zone 1 (ETP)
        configure_sip_log(f'{tarfile}/log.xml', str(sip_id), datetime.now().strftime("%Y-%m-%dT%H:%M:%S+02:00"))
        info_dict = gather_file_info(tarfile, os.path.basename(tarfile))
        info_dict.update(gather_file_info(content_path_label.cget("text"), f'{os.path.basename(tarfile)}/content'))
        configure_sip_premis(f'{tarfile}/administrative_metadata/premis.xml', str(sip_id), info_dict)
        configure_sip_mets(f'{tarfile}/mets.xml', str(sip_id), datetime.now().strftime("%Y-%m-%dT%H:%M:%S+02:00"), f'{tarfile}/administrative_metadata/premis.xml', info_dict)
        pack_sip(tarfile, sip_id, content_path_label.cget("text"))
        configure_sip_info(f'{output_folder}/info.xml', f'{tarfile}.tar', str(sip_id), datetime.now().strftime("%Y-%m-%dT%H:%M:%S+02:00"))

        #Zone 2 (ETA)
        aic_id = uuid1()
        os.rename(output_folder, f'./{aic_id}')
        configure_aic_log(f'./{aic_id}/{sip_id}/log.xml', str(aic_id), str(sip_id), datetime.now().strftime("%Y-%m-%dT%H:%M:%S+02:00"))
        
        PROGRESS_BAR.stop()
        logging("Process complete!")
        customtkinter.CTkButton(tabview.tab(3), text="Finish", command=lambda: sys.exit()).grid(row=5, column=0, columnspan=5, sticky="NSEW")
    except:
        with open("./error_log", "w", encoding="utf-8") as fo:
            traceback.print_exc(file=fo)

#A helper function for logging messages to the output textbox for the user to see
def logging(message: str):
    LOG_BOX.insert(tkinter.END, f'[{datetime.now().strftime("%d/%m/%y - %H:%M:%S")}]: {message}\n')
    window.update()

#A helper function which makes the relevant combo-box only list results relevant to what is already written in it
def combo_helper(element: customtkinter.CTkComboBox, list: list):
    element.configure(values=[i for i in list if element.get().lower() in i.lower()])

#A helper function for changing the user's username
def set_username():
    global USERNAME
    val = customtkinter.CTkInputDialog(title=' ', text='Set your username:').get_input()
    USERNAME = val if val else USERNAME
    menu.entryconfigure(1, label=f'Username: {USERNAME}')

#A helper function for initializing a grid for a specified widget
def configure_grid(colsize: int, rowsize: int, widget):
    for col in range(colsize):
        widget.columnconfigure(col, weight=1, pad=0)
    for row in range(rowsize):
        widget.rowconfigure(row, weight=1, pad=0)

#Initializes the main window
customtkinter.set_default_color_theme("blue")
customtkinter.set_appearance_mode("dark")
window = customtkinter.CTk()
window.title("Archive Package Creator")
window.geometry('{width}x{height}+{pos_right}+{pos_down}'.format(width=(window.winfo_screenwidth() // 2)+(window.winfo_screenwidth() // 3), height=(window.winfo_screenheight() // 2)+(window.winfo_screenheight() // 3), pos_right=(window.winfo_screenwidth() // 2)-((5*window.winfo_screenwidth()) // 12), pos_down=(window.winfo_screenheight() // 2)-((5*window.winfo_screenheight()) // 12)))

#Create tabview
tabview = customtkinter.CTkTabview(master=window, state="disabled")
for i in range(1,4):
    tabview.add(i)
tabview.pack(anchor=tkinter.CENTER, fill=tkinter.BOTH, expand=True, padx=10, pady=10)

#Configure tab 1
configure_grid(14,11,tabview.tab(1))

content_path_label = customtkinter.CTkLabel(tabview.tab(1), text="", fg_color="grey", corner_radius=8)
customtkinter.CTkButton(tabview.tab(1), text="Browse Content", command=lambda: browse_files(content_path_label)).grid(row=1, column=12, columnspan=1, sticky="NSEW")
descriptive_path_label = customtkinter.CTkLabel(tabview.tab(1), text="", fg_color="grey", corner_radius=8)
customtkinter.CTkButton(tabview.tab(1), text="Browse Descriptive Metadata (optional)", command=lambda: browse_files(descriptive_path_label)).grid(row=4, column=12, columnspan=1, sticky="NSEW")
administrative_path_label = customtkinter.CTkLabel(tabview.tab(1), text="", fg_color="grey", corner_radius=8)
customtkinter.CTkButton(tabview.tab(1), text="Browse Administrative Metadata (optional)", command=lambda: browse_files(administrative_path_label)).grid(row=7, column=12, columnspan=1, sticky="NSEW")
customtkinter.CTkButton(tabview.tab(1), text="Continue", command=lambda: tabview.set(2) if content_path_label.cget("text") else messagebox.showerror("Error", "No content path specified.")).grid(row=10, column=0, columnspan=14, sticky="NSEW")

content_path_label.grid(row=1, column=1, columnspan=9, sticky="NSEW")
descriptive_path_label.grid(row=4, column=1, columnspan=9, sticky="NSEW")
administrative_path_label.grid(row=7, column=1, columnspan=9, sticky="NSEW")

#Configure tab 2
configure_grid(7,8,tabview.tab(2))

TEXT_LIST = [StringVar(tabview.tab(2), name=f'{i}') for i in range(16)]
frame1 = customtkinter.CTkFrame(tabview.tab(2))
frame1.grid(row=1, column=1, columnspan=2, rowspan=2, sticky="NSEW")
configure_grid(6,8,frame1)
frame2 = customtkinter.CTkFrame(tabview.tab(2))
frame2.grid(row=1, column=4, columnspan=2, rowspan=2, sticky="NSEW")
configure_grid(6,6,frame2)
frame3 = customtkinter.CTkFrame(tabview.tab(2))
frame3.grid(row=4, column=1, columnspan=2, rowspan=2, sticky="NSEW")
configure_grid(6,5,frame3)
frame4 = customtkinter.CTkFrame(tabview.tab(2))
frame4.grid(row=4, column=4, columnspan=2, rowspan=2, sticky="NSEW")
configure_grid(6,4,frame4)
customtkinter.CTkButton(tabview.tab(2), text="Create Dias Package", command=lambda: threading.Thread(target=main_func, daemon=True).start() if all(len(i.get()) != 0 for i in TEXT_LIST) else messagebox.showerror("Error", "All input fields require input.")).grid(row=7, column=0, columnspan=7, sticky="NSEW")

customtkinter.CTkLabel(master=frame1, text="Label:", text_color="white", corner_radius=4, anchor='e').grid(row=1, column=1, sticky="EW")
label_entry = customtkinter.CTkEntry(master=frame1, textvariable=TEXT_LIST[4], corner_radius=4)
customtkinter.CTkLabel(master=frame1, text="System:", text_color="white", corner_radius=4, anchor='e').grid(row=2, column=1, sticky="EW")
system_combo = customtkinter.CTkComboBox(master=frame1, variable=TEXT_LIST[0], corner_radius=4, values=SYSTEM_LIST)
customtkinter.CTkLabel(master=frame1, text="System Version:", text_color="white", corner_radius=4, anchor='e').grid(row=3, column=1, sticky="EW")
system_ver_entry = customtkinter.CTkEntry(master=frame1, textvariable=TEXT_LIST[1], corner_radius=4)
customtkinter.CTkLabel(master=frame1, text="Submission Agreement:", text_color="white", corner_radius=4, anchor='e').grid(row=4, column=1, sticky="EW")
submission_entry = customtkinter.CTkEntry(master=frame1, textvariable=TEXT_LIST[2], corner_radius=4)
customtkinter.CTkLabel(master=frame1, text="Archivist System Type:", text_color="white", corner_radius=4, anchor='e').grid(row=5, column=1, sticky="EW")
type_combo = customtkinter.CTkComboBox(master=frame1, variable=TEXT_LIST[5], corner_radius=4, values=["SIARD", "NOARK-5", "Postjournaler", "Annet"])
customtkinter.CTkLabel(master=frame1, text="Period Start:", text_color="white", corner_radius=4, anchor='e').grid(row=6, column=1, sticky="EW")
period_start_entry = customtkinter.CTkEntry(master=frame1, textvariable=TEXT_LIST[10], corner_radius=4)
customtkinter.CTkLabel(master=frame1, text="Period End:", text_color="white", corner_radius=4, anchor='e').grid(row=6, column=3, sticky="EW")
period_end_entry = customtkinter.CTkEntry(master=frame1, textvariable=TEXT_LIST[11], corner_radius=4)
customtkinter.CTkLabel(master=frame2, text="Owner Organization:", text_color="white", corner_radius=4, anchor='e').grid(row=1, column=1, sticky="EW")
owner_org_combo = customtkinter.CTkComboBox(master=frame2, variable=TEXT_LIST[6], corner_radius=4, values=MUNICIPALITY_LIST)
customtkinter.CTkLabel(master=frame2, text="Archivist Organization:", text_color="white", corner_radius=4, anchor='e').grid(row=2, column=1, sticky="EW")
archivist_org_combo = customtkinter.CTkComboBox(master=frame2, variable=TEXT_LIST[3], corner_radius=4, values=MUNICIPALITY_LIST)
customtkinter.CTkLabel(master=frame2, text="Submitter Organization:", text_color="white", corner_radius=4, anchor='e').grid(row=3, column=1, sticky="EW")
submitter_org_combo = customtkinter.CTkComboBox(master=frame2, variable=TEXT_LIST[12], corner_radius=4, values=MUNICIPALITY_LIST)
customtkinter.CTkLabel(master=frame2, text="Submitter Person:", text_color="white", corner_radius=4, anchor='e').grid(row=4, column=1, sticky="EW")
submitter_pers_entry = customtkinter.CTkEntry(master=frame2, textvariable=TEXT_LIST[13], corner_radius=4)
customtkinter.CTkLabel(master=frame3, text="Producer Organization:", text_color="white", corner_radius=4, anchor='e').grid(row=1, column=1, sticky="EW")
producer_org_entry = customtkinter.CTkEntry(master=frame3, textvariable=TEXT_LIST[7], corner_radius=4)
customtkinter.CTkLabel(master=frame3, text="Producer Person:", text_color="white", corner_radius=4, anchor='e').grid(row=2, column=1, sticky="EW")
producer_pers_entry = customtkinter.CTkEntry(master=frame3, textvariable=TEXT_LIST[8], corner_radius=4)
customtkinter.CTkLabel(master=frame3, text="Producer Software:", text_color="white", corner_radius=4, anchor='e').grid(row=3, column=1, sticky="EW")
producer_software_entry = customtkinter.CTkEntry(master=frame3, textvariable=TEXT_LIST[9], corner_radius=4)
customtkinter.CTkLabel(master=frame4, text="Creator:", text_color="white", corner_radius=4, anchor='e').grid(row=1, column=1, sticky="EW")
creator_entry = customtkinter.CTkEntry(master=frame4, textvariable=TEXT_LIST[14], corner_radius=4)
customtkinter.CTkLabel(master=frame4, text="Preserver:", text_color="white", corner_radius=4, anchor='e').grid(row=2, column=1, sticky="EW")
preserver_entry = customtkinter.CTkEntry(master=frame4, textvariable=TEXT_LIST[15], corner_radius=4)

system_combo.bind('<KeyRelease>', lambda _: combo_helper(system_combo,SYSTEM_LIST))
type_combo.bind('<KeyRelease>', lambda _: combo_helper(type_combo,["SIARD", "NOARK-5", "Postjournaler"]))
owner_org_combo.bind('<KeyRelease>', lambda _: combo_helper(owner_org_combo,MUNICIPALITY_LIST))
archivist_org_combo.bind('<KeyRelease>', lambda _: combo_helper(archivist_org_combo,MUNICIPALITY_LIST))
submitter_org_combo.bind('<KeyRelease>', lambda _: combo_helper(submitter_org_combo,MUNICIPALITY_LIST))

label_entry.grid(row=1, column=2, columnspan=3, sticky="EW")
system_combo.grid(row=2, column=2, columnspan=3, sticky="EW")
system_ver_entry.grid(row=3, column=2, columnspan=3, sticky="EW")
submission_entry.grid(row=4, column=2, columnspan=3, sticky="EW")
type_combo.grid(row=5, column=2, columnspan=3, sticky="EW")
period_start_entry.grid(row=6, column=2, sticky="EW")
period_end_entry.grid(row=6, column=4, sticky="EW")
owner_org_combo.grid(row=1, column=2, columnspan=3, sticky="EW")
archivist_org_combo.grid(row=2, column=2, columnspan=3, sticky="EW")
submitter_org_combo.grid(row=3, column=2, columnspan=3, sticky="EW")
submitter_pers_entry.grid(row=4, column=2, columnspan=3, sticky="EW")
producer_org_entry.grid(row=1, column=2, columnspan=3, sticky="EW")
producer_pers_entry.grid(row=2, column=2, columnspan=3, sticky="EW")
producer_software_entry.grid(row=3, column=2, columnspan=3, sticky="EW")
creator_entry.grid(row=1, column=2, columnspan=3, sticky="EW")
preserver_entry.grid(row=2, column=2, columnspan=3, sticky="EW")

#Configure tab 3
configure_grid(5,6,tabview.tab(3))

PROGRESS_BAR = customtkinter.CTkProgressBar(tabview.tab(3), mode="indeterminate")
LOG_BOX = customtkinter.CTkTextbox(tabview.tab(3), wrap="none", font=("",20))

LOG_BOX.grid(row=1, column=1, columnspan=3, rowspan=3, sticky="NSEW")
PROGRESS_BAR.grid(row=4, column=1, columnspan=3, sticky="EW")

#Adds a menubar at the top of the window
menu = Menu(master=window)
menu.add_command(label=f'Username: {USERNAME}', command=set_username)
menu.add_command(label='Import mets.xml Metadata', command=lambda: import_metadata(filedialog.askopenfile(initialdir="./", title="Choose metadata file", filetypes=[("XML files", "*.xml")])))
window.config(menu=menu)

#Runs the main window and keeps it going
window.mainloop()