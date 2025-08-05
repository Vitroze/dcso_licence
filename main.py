import requests 
import csv
import json

SHEET_ID = "1G1cOitq2COtwK_IDeZc4UjFvnWYv5_s5IGgoPuvmfMA"
lGID = [
    "1360221607",  # Sheet1
    "0",            # Sheet2
]

sURL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&gid=GID"

tData0 = []
tData1 = []

tListType = ["Aucun"]
tListWeaponLicense = ["Aucun"]

for GID in lGID:
    sResponse = requests.get(sURL.replace("GID", GID))
    if sResponse.status_code == 200:
        with open(f"sheet_{GID}.csv", "w", newline='', encoding='utf-8') as csvfile:
            sValue = sResponse.text
            sValue = sValue.replace("vv vvvvvvvvv ", "")
            sValue = sValue.replace("p ", "")
            sValue = sValue.replace("ADMINISTRATIVE CONTROL SUBDIVISION INFORMATIONS PERSONNELLES NOM & Pénom", "RPName")
            sValue = sValue.replace("Date de naissance", "DateBirth")
            sValue = sValue.replace("Adresse de résidence", "ResidentialAddress")
            sValue = sValue.replace("INFORMATIONS SUR L'ARME Catégorie", "Type")
            sValue = sValue.replace("INFORMATIONS SUR LA LICENCE Catégorie de la licence ", "Type")
            sValue = sValue.replace("Modèle de l'arme", "NameWeapon")
            sValue = sValue.replace("Numéro de série", "SerialNumber")
            sValue = sValue.replace("F ", "")
            

            csv_reader = csv.DictReader(sValue.splitlines())

            if GID == "0":
                tData0 = [row for row in csv_reader]
                for row in tData0:
                    sType = "Aucun" if len(row['Type']) == 0 or row['Type'] == "" else row['Type']
                    sType = ' '.join(sType.split()).lower().capitalize()
                    if sType not in tListType:
                        tListType.append(sType)
            else:
                tData1 = {' '.join(row['RPName'].split()).lower(): row for row in csv_reader}

                for row in tData1.values():
                    sType = "Aucun" if len(row['Type']) == 0 or row['Type'] == "" else row['Type']
                    sType = ' '.join(sType.split()).lower().capitalize()

                    if sType not in tListWeaponLicense:
                        tListWeaponLicense.append(sType)

from tkinter import *

def center(DFrame, iW, iH):
    iX, iY = DFrame.winfo_screenwidth() / 2 - iW / 2, DFrame.winfo_screenheight() / 2 - iH / 2

    DFrame.geometry(f"{iW}x{iH}+{int(iX)}+{int(iY)}")

DFrame = None

def clear_panel(DFrame: Tk):
    for widget in DFrame.winfo_children():
        widget.destroy()

def main(sPage = None): 
    global DFrame

    if DFrame is not None:        
        DFrame.destroy()

    DFrame = Tk()

    DFrame.title("DCSO - Permis Port d'arme")
    DFrame.geometry(center(DFrame, 800, 500))
    DFrame.protocol("WM_DELETE_WINDOW", DFrame.quit)

    DLabel = Label(DFrame, text="Rechercher par numéro série, nom, prénom,\nadresse de résidence et date de naissance", font=("Arial", 8))
    DLabel.place(x=10, y=20)
    
    DTextSearch = Entry(DFrame)
    DTextSearch.insert(END, "")
    DTextSearch.place(width = 200, x=10, y=60)
    DTextSearch.bind("<Return>", lambda e: Search(DTextSearch.get()))

    DLabel2 = Label(DFrame, text="Type d'arme", font=("Arial", 8))
    DLabel2.place(x=240, y=40)
    sValueOptionValue = StringVar(value="Aucun")
    DComboBox = OptionMenu(DFrame, sValueOptionValue, *tListType)
    DComboBox.place(x=240, y=60, width=220)

    DLabel3 = Label(DFrame, text="Permis port d'arme", font=("Arial", 8))
    DLabel3.place(x=460, y=40)
    sValueOptionValueLicense = StringVar(value="Aucun")
    DComboBoxLicense = OptionMenu(DFrame, sValueOptionValueLicense, *tListWeaponLicense)
    DComboBoxLicense.place(x=460, y=60, width=220)
    sValueOptionValue.trace_add("write", lambda *args: Search(DTextSearch.get()))
    sValueOptionValueLicense.trace_add("write", lambda *args: Search(DTextSearch.get()))


    canvas = Canvas(DFrame, width=800, height=370)
    canvas.place(x=0, y=100)

    Scroll = Scrollbar(DFrame, orient=VERTICAL, command=canvas.yview)
    Scroll.place(x=780, y=100, height=400)

    scrollable_frame = Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=Scroll.set)

    print(len(tData0), len(tData1))
    def Search(sValue = None):
        clear_panel(scrollable_frame)
        i = 0

        print("Searching for:", sValue, sValueOptionValue.get())
        
        for tDataSheet in tData0:

            sRPName = ' '.join(tDataSheet['RPName'].split()).lower()
            if sValue is not None:

                sValue = sValue.strip()
                sValue = sValue.replace("-", "")
                sValue = ' '.join(sValue.split()).lower()

                sSerialNumber = tDataSheet['SerialNumber'].replace("-", "")

                sType = "Aucun" if len(tDataSheet['Type']) == 0 or tDataSheet['Type'] == "" else tDataSheet['Type']
                sType = ' '.join(sType.split()).lower().capitalize()

                if sValueOptionValue.get() != "Aucun" and sType != sValueOptionValue.get():
                    continue

                sTypeLicense = "Aucun" if len(tData1.get(sRPName, {}).get("Type", "Aucun")) == 0 or tData1.get(sRPName, {}).get("Type", "Aucun") == "" else tData1.get(sRPName, {}).get("Type", "Aucun")
                sTypeLicense = ' '.join(sTypeLicense.split()).lower().capitalize()

                if sValueOptionValueLicense.get() != "Aucun" and sTypeLicense != sValueOptionValueLicense.get():
                    continue

                if not (sValue in sSerialNumber.lower() or
                        sValue in sRPName or
                        sValue in tDataSheet['NameWeapon'].lower() or
                        sValue in tDataSheet['DateBirth'].lower() or
                        sValue in tDataSheet['ResidentialAddress'].lower()):
                    continue

            DPanel = Frame(scrollable_frame, width=760, height=40, bg="lightgrey")
            DPanel.grid(row=i, column=0, pady=5, padx=10)

            sPermis = tData1.get(sRPName, {}).get("Type", "Aucun")
            DLabel = Label(DPanel, text=f"{' '.join(tDataSheet['RPName'].split())} - {tDataSheet['NameWeapon']} ({tDataSheet['SerialNumber']}) - Permis : {sPermis}", font=("Arial", 12), bg="lightgrey")
            DLabel.place(x=10, y=10)

            if i >= 1000:
                break

            i += 1

    Search()
    DFrame.mainloop()

main()