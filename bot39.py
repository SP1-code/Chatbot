import easyocr
from PIL import Image, ImageOps, ImageEnhance
from io import BytesIO
import numpy as np
import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import re
from datetime import datetime
import csv  # Voor het opslaan van logs

# Initialiseer de EasyOCR-lezer
reader = easyocr.Reader(['nl', 'en', 'fr'])  # Ondersteuning voor Nederlands, Engels en Frans

# Laad Excel-bestand (meerdere tabbladen)
def load_excel_data(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        data = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
        return data
    except Exception as e:
        print(f"Fout bij het laden van Excel-bestand: {e}")
        return {}

# Haal de lijst van geautoriseerde gebruikers op uit het Excel-bestand
def load_authorized_users(file_path):
    try:
        users_data = pd.read_excel(file_path, sheet_name="gebruikers")
        authorized_users = set(users_data["gebruikersnaam"].str.lower())  # Veronderstel dat de kolom 'gebruikersnaam' heet
        return authorized_users
    except Exception as e:
        print(f"Fout bij het laden van de gebruikerslijst: {e}")
        return set()

# Verwijder specifieke kleuren uit de afbeelding
def remove_colors(image):
    image = image.convert("RGB")
    r, g, b = image.split()
    r = r.point(lambda i: i * 0.3)
    g = g.point(lambda i: i * 0.9)
    b = b.point(lambda i: i * 0.2)
    return Image.merge("RGB", (r, g, b))

# Verwerk afbeelding voor OCR
def preprocess_image(image):
    no_color_image = remove_colors(image)
    gray_image = no_color_image.convert("L")
    contrast_image = ImageEnhance.Contrast(gray_image).enhance(2.0)
    binary_image = contrast_image.point(lambda p: 255 if p > 180 else 0)
    return binary_image

# OCR-functie met EasyOCR
def extract_text_from_image(image_bytes):
    try:
        image = Image.open(BytesIO(image_bytes))
        processed_image = preprocess_image(image)

        # Converteer naar een numpy-array
        processed_array = np.array(processed_image)

        # EasyOCR tekstextractie
        result = reader.readtext(processed_array)
        text = "\n".join([text for _, text, _ in result])

        # Vervang "fgutenlijst" door "foutenlijst"
        text = text.replace("Fgutenlijst", "Foutenlijst")

        # Haal het eerste woord uit de tekst
        first_word = text.split()[0] if text else ""

        # Zoek naar "Foutenlijst" en de cijfers die daarop volgen
        foutcode_numbers = []
        match = re.search(r'Foutenlijst\s*(\d+)', text)
        if match:
            foutcode_numbers = match.group(1)

        # Toevoegen van de resultaten aan de tekst
        additional_info = f"{first_word} {foutcode_numbers}"

        # Voeg de extra informatie toe aan het uiteindelijke bericht
        text = additional_info

        return text
    except Exception as e:
        return f"Fout bij het uitvoeren van OCR: {e}"

# Valideer module en foutcodes
def validate_module_and_foutcodes(input_parts, data):
    modules = [module.lower() for module in data.keys()]
    module, foutcodes = None, []

    for part in input_parts:
        part_lower = part.lower()
        if part_lower in modules:
            module = next(original for original in data.keys() if original.lower() == part_lower)
        else:
            foutcodes.extend(part.lstrip('0').split())

    foutcodes = [code for item in foutcodes for code in re.split('[\s\+\-\.,]+', item) if code]
    return module, foutcodes

# Haal een antwoord op uit een specifieke module
def get_response(module, foutcode, data):
    if module not in data:
        return None, "Ongeldige module."

    module_data = data[module]
    if "Foutcode" not in module_data.columns:
        return None, "Module bevat geen foutcodes."

    match = module_data[module_data["Foutcode"].astype(str).str.lower() == foutcode.lower()]
    if not match.empty:
        beschrijving = match.iloc[0]["Foutcodebeschrijving"]
        oplossingen = [match.iloc[0][f"Oplossing {i}"] for i in range(1, 6)]
        return (beschrijving, oplossingen), None

    return None, "Ongeldige foutcode."

# Haal een antwoord op uit de tab "ALGEMEEN"
def get_general_response_by_keywords(vraag, data):
    algemeen_data = data.get("ALGEMEEN")
    if algemeen_data is not None and "Vraag" in algemeen_data.columns:
        vraag_woorden = set(vraag.lower().split())
        relevantie_scores = []

        for index, row in algemeen_data.iterrows():
            if pd.isna(row["Vraag"]):
                continue
            vraag_in_algemeen = row["Vraag"].lower()
            vraag_in_algemeen_woorden = set(vraag_in_algemeen.split())
            overlap = vraag_woorden & vraag_in_algemeen_woorden
            relevantie_scores.append((len(overlap), index, row["Antwoord"]))

        relevantie_scores.sort(reverse=True, key=lambda x: x[0])
        if relevantie_scores and relevantie_scores[0][0] > 0:
            _, _, antwoord = relevantie_scores[0]
            return antwoord

    return None

# Opslaan van gebruikersinteracties in een CSV-bestand
def log_user_interaction(user_id, username, question, response):
    log_file = "user_logs.csv"
    try:
        with open(log_file, mode="a", newline='', encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow([datetime.now(), user_id, username, question, response])
    except Exception as e:
        print(f"Fout bij het loggen van de interactie: {e}")

# Startcommando
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Welkom! Upload een foto met tekst of stel een foutcodevraag.")

# Verwerk tekstberichten
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_message = update.message.text
    user_id = update.message.from_user.id
    username = update.message.from_user.username or "Onbekend"

    # Controleer of de gebruiker geautoriseerd is
    if username.lower() not in authorized_users:
        await update.message.reply_text("User not authorized")
        return  # Stop verder verwerken als de gebruiker niet geautoriseerd is

    input_parts = user_message.lower().split()
    module, foutcodes = validate_module_and_foutcodes(input_parts, data)

    if module:
        for foutcode in foutcodes:
            response, error = get_response(module, foutcode, data)
            if response:
                beschrijving, oplossingen = response
                oplossingen_text = "\n".join(
                    f"{i + 1}. {oplossing}" for i, oplossing in enumerate(oplossingen) if pd.notna(oplossing))
                antwoord = (f"Module: {module}\nFoutcode: {foutcode}\nBeschrijving: {beschrijving}\n\n"
                            f"Oplossingen:\n{oplossingen_text}")
                await update.message.reply_text(antwoord)
                log_user_interaction(user_id, username, user_message, antwoord)
            else:
                error_message = f"Foutcode {foutcode}: {error}"
                await update.message.reply_text(error_message)
                log_user_interaction(user_id, username, user_message, error_message)
    else:
        vraag = " ".join(input_parts)
        antwoord = get_general_response_by_keywords(vraag, data)

        if antwoord:
            await update.message.reply_text(f"Vraag: {vraag}\nAntwoord: {antwoord}")
            log_user_interaction(user_id, username, user_message, antwoord)
        else:
            no_info_message = ("Geen module of bijbehorende foutcode gevonden. "
                               "Ook geen overeenkomende vraag in 'ALGEMEEN'.")
            await update.message.reply_text(no_info_message)
            log_user_interaction(user_id, username, user_message, no_info_message)

# Verwerk foto's
async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    photo = update.message.photo[-1]
    photo_file = await photo.get_file()
    photo_bytes = await photo_file.download_as_bytearray()

    user_id = update.message.from_user.id
    username = update.message.from_user.username or "Onbekend"

    # Controleer of de gebruiker geautoriseerd is
    if username.lower() not in authorized_users:
        await update.message.reply_text("User not authorized")
        return  # Stop verder verwerken als de gebruiker niet geautoriseerd is

    extracted_text = extract_text_from_image(photo_bytes)
    await update.message.reply_text(f"Uit de afbeelding geëxtraheerde tekst:\n{extracted_text}")
    log_user_interaction(user_id, username, "Foto geüpload", extracted_text)

    input_parts = extracted_text.lower().split()
    module, foutcodes = validate_module_and_foutcodes(input_parts, data)

    if module:
        for foutcode in foutcodes:
            response, error = get_response(module, foutcode, data)
            if response:
                beschrijving, oplossingen = response
                oplossingen_text = "\n".join(
                    f"{i + 1}. {oplossing}" for i, oplossing in enumerate(oplossingen) if pd.notna(oplossing))
                antwoord = (f"Module: {module}\nFoutcode: {foutcode}\nBeschrijving: {beschrijving}\n\n"
                            f"Oplossingen:\n{oplossingen_text}")
                await update.message.reply_text(antwoord)
                log_user_interaction(user_id, username, extracted_text, antwoord)
            else:
                error_message = f"Foutcode {foutcode}: {error}"
                await update.message.reply_text(error_message)
                log_user_interaction(user_id, username, extracted_text, error_message)
    else:
        vraag = extracted_text
        antwoord = get_general_response_by_keywords(vraag, data)

        if antwoord:
            await update.message.reply_text(f"Vraag: {vraag}\nAntwoord: {antwoord}")
            log_user_interaction(user_id, username, vraag, antwoord)
        else:
            no_info_message = "Geen relevante informatie gevonden op basis van de afbeelding."
            await update.message.reply_text(no_info_message)
            log_user_interaction(user_id, username, vraag, no_info_message)

# Start de bot
def main():
    TOKEN = "8022858820:AAFytE3YsPYjt11df6wGAXMuGQvQowRTRDM"
    global data, authorized_users
    data = load_excel_data("modules_met_uitleg.xlsx")
    authorized_users = load_authorized_users("modules_met_uitleg.xlsx")  # Voeg deze lijn toe

    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))

    print("Bot is gestart...")
    app.run_polling()

if __name__ == "__main__":
    main()
