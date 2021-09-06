# Importing the required packages
import os
#import re
import math
import uuid
import datetime as dt
import pandas as pd
import numpy as np
import lxml.html as lh
from lxml import etree as et
from lxml.builder import E


# Setting the current working directory
path = r'C:\Users\m\Documents\PROGRAMMING SCRIPTS\CLASH MATRIX AUTOMATION\Clash Matrix Data'
os.chdir(path)
os.getcwd()

# The code below generates a string format of the current date and time.
now = dt.datetime.now().strftime("%Y-%m-%d") 

FILE_NAME = f"{now}-CLASH TEST (ALL DISCIPLINES)"

# DataFrame from clash detection matrix
cdm = pd.read_excel(
    r'E:\NAVISWORKS CLASH TEST FILES\Clash Matrix Data\Navisworks_Clash_Matrix.xlsx', 
    sheet_name=0, 
    index_col=[0,1,2], 
    header=[0,1,2],
    keep_default_na=True
)

cdm.head()

cdm.info()

cdm_val = cdm.values.tolist()
cdm_idx = cdm.index.to_list() 

print(cdm_idx)
clash_tests_names = []
clash_tols_types = [] 
rgt_loc = []
lft_loc = []

for idx_r, row in enumerate(cdm_val):
    for idx_c, clash_tol_type in enumerate(row):
        if (clash_tol_type is not np.nan) and (type(clash_tol_type) is str):
            clash_grp_c = cdm_idx[idx_c][2] # Clash group from column
            clash_grp_r = cdm_idx[idx_r][2] # Clash group from row
            clash_test = clash_grp_r + ' _VS_ ' + clash_grp_c
            # Left locator construction
            lft_con = f"lcop_selection_set_tree/{cdm_idx[idx_r][0]}/{cdm_idx[idx_r][1]}-{clash_grp_r}"
            # Right locator construction
            rgt_con = f"lcop_selection_set_tree/{cdm_idx[idx_c][0]}/{cdm_idx[idx_c][1]}-{clash_grp_c}"
            # Appending to lists
            clash_tests_names.append(clash_test)
            clash_tols_types.append(clash_tol_type)
            rgt_loc.append(rgt_con)
            lft_loc.append(lft_con)
            # print(f'Clash Group: {clash_grp_r}, Type & Tolerance: {clash_tols_types}, Clash Test: {clash_test}')
        else:
            pass

#print(clash_tests_names)

clash_types = []  # clash type
clash_tols = [(int(i[2:]) / 1000) for i in clash_tols_types]  # clash tolerance

for i in clash_tols_types:
    if i[0].lower() == 'c':
        clash_types.append("clearance")
    elif i[0].lower() == 'h':
        clash_types.append("hard")
    else:
        pass

Element_category = [cdm_idx[idx][1] for idx, row in enumerate(cdm_idx)]  # Clash priority list
Element_type = [cdm_idx[idx][2] for idx, row in enumerate(cdm_idx)]  # Clash group list
disciplines = [cdm_idx[idx][0] for idx, row in enumerate(cdm_idx)]  # Discipline list

dfSS = pd.DataFrame({
    'Element Category': Element_category,
    'Element Type': Element_type,
    'Discipline': disciplines,
})

dfSS.head()

# getting unique values for clash priority and discipline

cPrLi_unq = dfSS['Element Category'].unique().tolist()  # unique clash priority
dispLi_unq = dfSS['Discipline'].unique().tolist()  # unique clash discipline

# Setting the xml schema and generating the root

qnmAtt = et.QName("http://www.w3.org/2001/XMLSchema-instance", "noNamespaceSchemaLocation")

ct_root = et.Element(
    'exchange',
    {qnmAtt: "http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd"},
    units="m"
)

baTest = et.Element(
    'batchtest', {
        "name": FILE_NAME, 
        "internal_name": FILE_NAME, 
        "units": "m"
        } 
)

ct_root.append(baTest)

clashtests = et.SubElement(baTest, "clashtests")

# Generating clash test xml
for idx_cT, cT in enumerate(clash_tests_names):
    tT = clash_types[idx_cT]  # test type
    cTl = clash_tols[idx_cT]  # tolerance
    rLoc = rgt_loc[idx_cT]  # right locator
    lLoc = lft_loc[idx_cT]  # left locator
    clashtest = et.SubElement(
        clashtests, "clashtest", {
            "name": str(cT),
            "test_type": str(tT),
            "status": "new",
            "tolerance": str(cTl),
            "merge_composites": "0"
        }
    )

    linkage = et.SubElement(
        clashtest, "linkage", {
            "mode": "none"
        }
    )  

    left = et.SubElement(clashtest, "left")
    # left children
    clashselectionL = et.SubElement(
        left, "clashselection", {
            "selfintersect": "0",
            "primtypes": "1" 
        }
    )
    # clashselection child (left locator)
    lLocator = et.SubElement(clashselectionL, "locator")
    lLocator.text = lLoc

    right = et.SubElement(clashtest, "right")  
    # right children
    clashselectionR = et.SubElement(
        right, "clashselection", {
            "selfintersect": "0",
            "primtypes": "1" 
        }
    )
    # clashselection child (right locator)
    rLocator = et.SubElement(clashselectionR, "locator")
    rLocator.text = rLoc

    rules = et.SubElement(clashtest, "rules")

print(len(ct_root[0][0]))

print(et.tostring(ct_root, pretty_print=True, encoding="UTF-8", xml_declaration=True).decode('utf-8'))

selectionsets = et.SubElement(baTest, 'selectionsets')

# Creating and appending all viewfolders based on discipline, to selectionsets
for disp in dispLi_unq:
    viewfolder = et.SubElement(
        selectionsets, 'viewfolder', {
            "name": str(disp),
            "guid": str(uuid.uuid4())
        }
    )

# Generating all search sets
for disp in dispLi_unq:
    vwFr = ct_root.find(f'.//viewfolder[@name="{disp}"]')
    for row in dfSS.itertuples(index=False, name=None):  # enumerate not required
        if row[2] == disp:
            # Constructing the selectionset name
            ssCon = [row[0], row[1]]
            ssName = '-'.join(ssCon)
            # Creating and appending selectionsets 
            selectionset = et.SubElement(
                vwFr, 'selectionset', {
                    "name": str(ssName),
                    "guid": str(uuid.uuid4())
                }
            )
            
            findspec = et.SubElement(
                selectionset, 'findspec', {
                    "mode": "all",
                    "disjoint": "0"
                }
            )

            conditions = et.SubElement(findspec, 'conditions')

            locator = et.SubElement(findspec, 'locator')
            locator.text = "/"

            # conditions children
            condition1 = et.SubElement(
                conditions, 'condition', {
                    "test": "contains",
                    "flags": "10"
                }
            )

            category1 = et.SubElement(condition1, 'category')

            nmCt1 = et.SubElement(
                category1, 'name', {
                    "internal": "AecDbPropertySet"
                }
            ) 
            nmCt1.text = "Element"

            prop1 = et.SubElement(condition1, 'property')

            nmPp1 = et.SubElement(
                prop1, 'name', {
                    "internal": "1"
                }
            ) 
            nmPp1.text = "Category"

            value1 = et.SubElement(condition1, 'value')

            data1 = et.SubElement(
                value1, "data", {
                    "type": "wstring"
                }
            )
            data1.text = str(row[0])

            # second condition
            condition2 = et.SubElement(
                conditions, 'condition', {
                    "test": "contains",
                    "flags": "10"
                }
            )

            category2 = et.SubElement(condition2, 'category')

            nmCt2 = et.SubElement(
                category2, 'name', {
                    "internal": "AecDbPropertySet"
                }
            ) 
            nmCt2.text = "Element"

            prop2 = et.SubElement(condition2, 'property')

            nmPp2 = et.SubElement(
                prop2, 'name', {
                    "internal": "3"
                }
            ) 
            nmPp2.text = "Type"

            value2 = et.SubElement(condition2, 'value')

            data2 = et.SubElement(
                value2, "data", {
                    "type": "wstring"
                }
            )
            data2.text = str(row[1])
        else:
            pass

# code below searches for the first viewfolder then shows its xml structure
vwFr_all = ct_root.find('.//viewfolder')
print(et.tostring(vwFr_all, pretty_print=True).decode('utf-8'))

ctData = et.tostring(ct_root, pretty_print=True, encoding="UTF-8", xml_declaration=True)

#Create .XML File and export for use as Clash Test in Navisworks

with open(f"{FILE_NAME}.xml", "wb") as nfXML:
    nfXML.write(ctData)