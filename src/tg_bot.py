import json
import telegram

with open("credentials/telegram_token.json") as file:
    data = json.load(file)
token = data["token"]

bot = telegram.Bot(token=token)

async def send_message(text, chat_id):
    async with bot:
        await bot.send_message(text=text, chat_id=chat_id, parse_mode = telegram.constants.ParseMode.HTML)

async def publish_empty_agenda(agenda_link, attachment_folder, meeting_number, meeting_date, meeting_time):
    msg = f"Kokous {meeting_number}/2025\n\nSeuraava kokous on keskiviikkona {meeting_date} klo {meeting_time}.\n\nEsityslista löytyy <a href='{agenda_link}'>täältä</a> ja se tulee täyttää klo 15 mennessä tiistaina. Kokouksessa käytävät liitteet tulee lisätä kansioon <a href='https://drive.google.com/drive/u/0/folders/{attachment_folder}'>{meeting_number}/2025</a>.\n\nEstyneisyydestä tulee ilmoittaa klo 15 mennessä tiistaina <a href='https://forms.gle/ztX5DYz5gtBKBXhJ6'>tällä lomakkeella</a>."
    print(msg)
    await send_message(msg, 2079819344)