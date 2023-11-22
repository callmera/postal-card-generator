import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display

# Define a class to represent each row
class Employee:
    def __init__(self, name, lastname, job_position, organization):
        self.name = name
        self.lastname = lastname
        self.job_position = job_position
        self.organization = organization

# Read data from Excel file
df = pd.read_excel('excel/sheet.xlsx')

# List to store class instances
employees = []

for index, row in df.iterrows():
    employee = Employee(row['نام'], row['نام خانوادگی'], row['سمت شغلی'], row['سازمان'])
    employees.append(employee)

def add_text_to_postal_card(template_path, employee, output_path):
    # Open the image
    image = Image.open(template_path)
    draw = ImageDraw.Draw(image)

    # Define font (use a .ttf file you have, or a default one)
    try:
        font = ImageFont.truetype("font/BYekan+ Bold.ttf", size=60)
    except IOError:
        font = ImageFont.load_default()
    
    try:
        font_2 = ImageFont.truetype("font/BYekan+ Bold.ttf", size=40)
    except IOError:
        font_2 = ImageFont.load_default()

    

    # Function to reshape and reorder Persian text
    def reshape_text(text):
        reshaped_text = arabic_reshaper.reshape(text)
        return get_display(reshaped_text)
    
    def center_text(height, text, alligner):
        text_width = len(text) * alligner
        x = image.width / 2 - text_width / 2
        y = height
        return (x, y)

    name_coords = center_text(352, employee.name + ' ' + employee.lastname, 22)
    job_position_coords = center_text(450, employee.job_position, 17)
    organization_coords = center_text(497, employee.organization, 17)

    # Add text to image
    draw.text(name_coords, reshape_text(employee.name + ' ' + employee.lastname), fill="white", font=font)
    draw.text(job_position_coords, reshape_text(employee.job_position), fill="white", font=font_2)
    draw.text(organization_coords, reshape_text(employee.organization), fill="white", font=font_2)

    # Save the image
    image.save(output_path)

# instance = employees[8]
# add_text_to_postal_card('postal/postal.jpg', instance, f'postal/{instance.name}-{instance.lastname}.jpg')
not_applied = []
for empl in employees:
    
    try:
        add_text_to_postal_card('postal/postal.jpg', empl, f'postal/result/{empl.name}-{empl.lastname}.jpg')
    except Exception as e:
        not_applied.append(empl.name + ' ' + empl.lastname)

for n in not_applied:
    print(f'not applied for {n}')
