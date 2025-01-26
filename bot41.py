import pandas as pd
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import re
from datetime import datetime
import csv  # Voor het opslaan van logs

# Laad Excel-bestand (meerdere tabbladen)
def load_excel_data(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        data = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
        return data
    except Exception as e:
        print(f"Fout bij het laden van Excel-bestand: {e}")
        return {}

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
def log_interaction(question, response):
    log_file = "user_logs.csv"
    try:
        with open(log_file, mode="a", newline='', encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow([datetime.now(), question, response])
    except Exception as e:
        print(f"Fout bij het loggen van de interactie: {e}")

# Startcommando
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Welkom! Stel je vraag met betrekking tot modules of foutcodes.")

# Verwerk tekstberichten
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_message = update.message.text

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
                log_interaction(user_message, antwoord)
            else:
                error_message = f"Foutcode {foutcode}: {error}"
                await update.message.reply_text(error_message)
                log_interaction(user_message, error_message)
    else:
        vraag = " ".join(input_parts)
        antwoord = get_general_response_by_keywords(vraag, data)

        if antwoord:
            await update.message.reply_text(f"Vraag: {vraag}\nAntwoord: {antwoord}")
            log_interaction(user_message, antwoord)
        else:
            no_info_message = ("Geen module of bijbehorende foutcode gevonden. "
                               "Ook geen overeenkomende vraag in 'ALGEMEEN'.")
            await update.message.reply_text(no_info_message)
            log_interaction(user_message, no_info_message)

# Start de bot
def main():
    TOKEN = "8022858820:AAFytE3YsPYjt11df6wGAXMuGQvQowRTRDM"
    global data
    data = load_excel_data("modules_met_uitleg.xlsx")

    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    print("Bot is gestart...")
    app.run_polling()

if __name__ == "__main__":
    main()
