import json
import telegram

with open("credentials/telegram_token.json") as file:
    data = json.load(file)
token = data["token"]

bot = telegram.Bot(token=token)

async def send_message(text, chat_id):
    async with bot:
        await bot.send_message(text=text, chat_id=chat_id, parse_mode = telegram.constants.ParseMode.HTML)

async def publish_empty_agenda(agenda_link, attachment_folder, meeting_number, meeting_date, meeting_time, inhibited):
    if inhibited:
        msg = (
            f"Kokous {meeting_number}/2025\n\n"
            f"Seuraava kokous on tiistaina {meeting_date} klo {meeting_time}.\n\n"
            f"Esityslista löytyy <a href='{agenda_link}'>täältä</a> ja se tulee täyttää klo 15 mennessä maanantaina. "
            f"Kokouksessa käytävät liitteet tulee lisätä kansioon "
            f"<a href='https://drive.google.com/drive/u/0/folders/{attachment_folder}'>{meeting_number}/2025</a>.\n\n"
            f"Estyneisyydestä tulee ilmoittaa klo 15 mennessä maanantaina "
            f"<a href='https://forms.gle/sK6JxFYH2m2h33zFA'>tällä lomakkeella</a>."
        )
    else:
        msg = (
            f"Kokous {meeting_number}/2025\n\n"
            f"Seuraava kokous on tiistaina {meeting_date} klo {meeting_time}.\n\n"
            f"Esityslista löytyy <a href='{agenda_link}'>täältä</a> ja se tulee täyttää klo 15 mennessä maanantaina. "
            f"Kokouksessa käytävät liitteet tulee lisätä kansioon "
            f"<a href='https://drive.google.com/drive/u/0/folders/{attachment_folder}'>{meeting_number}/2025</a>.\n\n"
            f"Estyneisyydestä ei tarvitse ilmoittaa."
        )
    print(msg)
    await send_message(msg, -1002461541855) # KIKHT25-tiedotus: -1002461541855, vain minä: 2079819344