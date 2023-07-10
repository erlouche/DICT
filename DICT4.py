import pdfplumber
import re
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import zipfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from tkinter import simpledialog


def extract_data_from_xml(xml_string):
    contact_matches = re.findall(r'<t:contact>(.*?)</t:contact>', xml_string, re.DOTALL)

    data = []
    for contact_match in contact_matches:
        societe_match = re.search(r'<t:societe>(.*?)</t:societe>', contact_match, re.DOTALL)
        courriel_match = re.search(r'<t:courriel>(.*?)</t:courriel>', contact_match, re.DOTALL)
        gere_fichiers_match = re.search(r'<t:gereLesFichiersDematerialises>(.*?)</t:gereLesFichiersDematerialises>', contact_match, re.DOTALL)
        format_fichiers_match = re.search(r'<t:formatDesFichiersDematerialises>(.*?)</t:formatDesFichiersDematerialises>', contact_match, re.DOTALL)

        societe_data = societe_match.group(1) if societe_match else None
        courriel_data = courriel_match.group(1) if courriel_match else None
        gere_fichiers_data = gere_fichiers_match.group(1) if gere_fichiers_match else None
        format_fichiers_data = format_fichiers_match.group(1) if format_fichiers_match else None

        data.append((societe_data, courriel_data, gere_fichiers_data, format_fichiers_data))

    return data

# Étape 1: Sélectionner le fichier ZIP contenant les fichiers
root = tk.Tk()
root.withdraw()
zip_file_path = filedialog.askopenfilename(filetypes=[("ZIP Files", "*.zip")])
dt = os.path.basename(zip_file_path)
zip_directory = os.path.dirname(os.path.abspath(zip_file_path))

# Étape 2: Extraire les fichiesr PDF du fichier ZIP

with zipfile.ZipFile(zip_file_path, 'r') as zip_file:
    for file_name in zip_file.namelist():
        if file_name.endswith(".pdf"):
            zip_file.extract(file_name, zip_directory)

pdf_resume = f"{zip_file_path[:-4]}_resume.pdf"

# Étape 3: Extraire le fichier XML se terminant par "description.xml" du fichier ZIP
xml_file_path = ""
with zipfile.ZipFile(zip_file_path, 'r') as zip_file:
    description_zip = None
    for file_name in zip_file.namelist():
        if file_name.endswith("description.zip"):
            description_zip = file_name
            break

    if description_zip:
        zip_file.extract(description_zip, "./")
        with zipfile.ZipFile(description_zip, 'r') as desc_zip:
            for file_name in desc_zip.namelist():
                if file_name.endswith("description.xml"):
                    desc_zip.extract(file_name, zip_directory)
                    xml_file_path = file_name
                    break


# Étape 4: Créer un DataFrame à partir du Résumé PDF
pattern = r"(\d+) ([A-Z][a-zA-Zéàè\s-]+)"
with pdfplumber.open(pdf_resume) as pdf:
    num_pages = len(pdf.pages)
    found_lines = []
    expected_number = 1
    for page_num in range(num_pages):
        page = pdf.pages[page_num]
        lines = page.extract_text().split("\n")

        for line in lines:
            match = re.search(pattern, line)
            if match:
                line_number = int(match.group(1))
                if line_number == expected_number:
                    found_lines.append(line)
                    expected_number += 1
                elif line_number > expected_number:
                    break
pdf_df = pd.DataFrame(found_lines, columns=['Société'])
pdf_df['Société'] = pdf_df['Société'].str.replace(r"^\d+\s", "")

# Étape 5: Créer un DataFrame à partir du XML
with open(xml_file_path, "r", encoding="utf-8") as file:
    xml_data = file.read()
    xml_df = pd.DataFrame(extract_data_from_xml(xml_data), columns=["Société", "Courriel", "Gère les fichiers dématérialisés", "Format des fichiers dématérialisés"])

# Étape 6: Créer un DataFrame combiné (pour garder l'ordre des cerfa)
combined_df = pdf_df.copy()
combined_df["Courriel"] = ""
combined_df["Gère les fichiers dématérialisés"] = ""
combined_df["Format des fichiers dématérialisés"] = ""

for index, row in pdf_df.iterrows():
    company_name = row["Société"]
    matching_row = xml_df[xml_df["Société"] == company_name]
    if not matching_row.empty:
        combined_df.loc[index, "Courriel"] = matching_row.iloc[0]["Courriel"]
        combined_df.loc[index, "Gère les fichiers dématérialisés"] = matching_row.iloc[0]["Gère les fichiers dématérialisés"]
        combined_df.loc[index, "Format des fichiers dématérialisés"] = matching_row.iloc[0]["Format des fichiers dématérialisés"]

print(combined_df)

# Étape 7: Envoyer un e-mail avec les fichiers joints
smtp_server = "smtp.gmail.com"
smtp_port = 587
default_sender_email = "eeun.auto@gmail.com"
sender_email = simpledialog.askstring("Adresse e-mail de l'expéditeur", "Veuillez entrer l'adresse e-mail de l'expéditeur :", initialvalue=default_sender_email)
sender_password = simpledialog.askstring("Mot de passe de l'expédieteur", "Mot de passe d'application (à générer dans gmail):", show='*')

def send_email(subject, body, to_email, attachments):
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    for attachment in attachments:
        pdf_path = os.path.join(zip_directory, attachment)
        with open(pdf_path, "rb") as file:  
            part = MIMEApplication(file.read(), Name=attachment)
        part["Content-Disposition"] = f'attachment; filename="{attachment}"'
        msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        print(f"Email sent to: {to_email}")
    except Exception as e:
        print(f"Error sending email to {to_email}: {str(e)}")

import tkinter.messagebox as messagebox

for index, row in combined_df.iterrows():
    if row["Gère les fichiers dématérialisés"] == "true" and row["Format des fichiers dématérialisés"] == "XML_PDF":
        company_name = row["Société"]
        email = row["Courriel"]
        zip_file_name = os.path.basename(zip_file_path)[:-4]  # Nom du fichier ZIP sans l'extension
        pdf_attachment = f"{zip_file_name}_emprise.pdf"
        pdf_cerfa = f"{zip_file_name}_{index + 1}.pdf"
        xml_attachment = xml_file_path
        
       
        print(f"Société: {company_name}")
        print(f"Courriel: {email}")
        print(f"Attachments: {pdf_attachment}, {xml_attachment}, {pdf_cerfa}")

        confirmation = messagebox.askyesno("Confirmation d'envoi", "Êtes-vous sûr de vouloir envoyer les e-mails ?")

        if not confirmation:
            print("Envoi d'e-mail annulé.")
            continue  

        attachments = [pdf_attachment, xml_attachment, pdf_cerfa]
        subject = f"{dt[:-4]}"
        body = f"Veuillez trouver ci-joint les pièces pour la DT"
        send_email(subject, body, email, attachments)
        
    elif row["Gère les fichiers dématérialisés"] == "false":
        company_name = row["Société"]
        messagebox.showinfo("Information", f"La société {company_name} ne gère pas les fichiers dématérialisés.")
        continue  

    elif row["Gère les fichiers dématérialisés"] == "true" and row["Format des fichiers dématérialisés"] == "XML":
        company_name = row["Société"]
        email = row["Courriel"]
        xml_attachment = xml_file_path

        print(f"Société: {company_name}")
        print(f"Courriel: {email}")
        print(f"Attachments: {xml_attachment}")

        confirmation = messagebox.askyesno("Confirmation d'envoi", "Êtes-vous sûr de vouloir envoyer les e-mails ?")

        if not confirmation:
            print("Envoi d'e-mail annulé.")
            continue  

        attachments = [xml_attachment]
        subject = f"{dt[:-4]}"
        body = f"Veuillez trouver ci-joint les pièces pour la DT"
        send_email(subject, body, email, attachments)


    
