import hashlib
import time
import requests
import urllib3
from tkinter import Tk, Label, Entry, Button, Canvas, StringVar, Toplevel, messagebox
from PIL import Image, ImageTk
from io import BytesIO
from requests.packages.urllib3.contrib.pyopenssl import inject_into_urllib3
import pandas as pd

# Disable SSL warnings and use pyOpenSSL's implementation
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
inject_into_urllib3()

PUBLIC_KEY = 'your public key'
PRIVATE_KEY = 'your private key'

class MarvelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Marvel Character Comics Downloader")

        self.label = Label(root, text="Enter the name of the Marvel character:")
        self.label.pack(pady=10)

        self.entry = Entry(root)
        self.entry.pack(pady=10)

        self.button = Button(root, text="Fetch Comics", command=self.fetch_comics)
        self.button.pack(pady=10)

        self.canvas = Canvas(root)
        self.canvas.pack(pady=10)

        self.comic_title = StringVar()
        self.label_comic_title = Label(root, textvariable=self.comic_title)
        self.label_comic_title.pack(pady=10)

        self.status = StringVar()
        self.label_status = Label(root, textvariable=self.status, fg="red")
        self.label_status.pack(pady=10)

        self.countdown_var = StringVar()
        self.label_countdown = Label(root, textvariable=self.countdown_var)
        self.countdown_var.set("")  # Initialize with an empty string
        self.label_countdown.pack(pady=10)

    def fetch_comics(self):
        character_name = self.entry.get()
        timestamp = str(time.time())
        hash_value = hashlib.md5((timestamp + PRIVATE_KEY + PUBLIC_KEY).encode('utf-8')).hexdigest()
        base_url = "https://gateway.marvel.com:443/v1/public/characters"
        params = {
            'name': character_name,
            'apikey': PUBLIC_KEY,
            'ts': timestamp,
            'hash': hash_value
        }

        response = requests.get(base_url, params=params, verify=False)
        data = response.json()

        if 'data' in data and 'results' in data['data'] and data['data']['results']:
            character_data = data['data']['results'][0]
            if 'comics' in character_data and 'available' in character_data['comics']:
                self.status.set("Downloading comic covers...")

                # Create a new Excel writer
                publisher_name = "Marvel"
                filename = f"{publisher_name}_{character_name.capitalize()}.xlsx"

                # Create a Pandas DataFrame to store data
                comic_titles = []
                comic_numbers = []
                comic_images = []

                # Fetch all comics by handling pagination
                offset = 0
                total_comics = character_data['comics']['available']
                while offset < total_comics:
                    params['offset'] = offset
                    response = requests.get(base_url, params=params, verify=False)
                    character_data = response.json().get('data', {}).get('results', [{}])[0]

                    if 'comics' in character_data:
                        for comic in character_data['comics'].get('items', []):
                            comic_title = comic.get('name', 'N/A')
                            comic_titles.append(comic_title)

                            # Split the title into title and number (if possible)
                            split_title = comic_title.rsplit('#', 1)
                            if len(split_title) == 2:
                                title, number = split_title
                                comic_numbers.append(number)
                            else:
                                title = split_title[0]
                                comic_numbers.append("N/A")

                            # Update GUI with current comic info
                            self.comic_title.set(title)
                            self.canvas.update()

                            # For simplicity, let's just get the first comic's image.
                            image_url = character_data.get('thumbnail', {}).get('path', '') + "." + character_data.get('thumbnail', {}).get('extension', '')
                            response = requests.get(image_url)
                            img_data = response.content
                            img = Image.open(BytesIO(img_data))
                            img = img.resize((200, 300))
                            photo = ImageTk.PhotoImage(img)
                            self.canvas.config(width=200, height=300)
                            self.canvas.create_image(100, 150, image=photo)
                            self.canvas.image = photo
                            self.canvas.update()

                            # Append comic image data
                            comic_images.append(img_data)

                    offset += 20  # Increase offset for pagination

                    # Display countdown
                    for seconds_left in range(30, -1, -1):
                        self.countdown_var.set(f"Next request in {seconds_left} seconds")
                        self.label_countdown.update()
                        time.sleep(1)

                # Create a Pandas DataFrame
                df = pd.DataFrame({
                    'Title': comic_titles,
                    'Number': comic_numbers,
                })

                # Create an ExcelWriter object
                writer = pd.ExcelWriter(filename, engine='xlsxwriter')

                # Convert the DataFrame to an XlsxWriter Excel object
                df.to_excel(writer, sheet_name=character_name.capitalize(), index=False)

                # Get the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets[character_name.capitalize()]

                # Add the comic images to the Excel file
                for idx, img_data in enumerate(comic_images, start=1):
                    image_stream = BytesIO(img_data)
                    worksheet.insert_image(f'C{idx + 1}', '', {'image_data': image_stream})

                # Close the ExcelWriter object
                writer.save()
                self.status.set(f"Download complete! Saved as {filename}")
            else:
                self.status.set("No comics found for this character.")
        else:
            self.status.set("Character not found. Please check the name.")

app = Tk()
MarvelApp(app)
app.mainloop()
