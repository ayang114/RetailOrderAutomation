import os
import pandas as pd
from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

app = Flask(__name__)

# Directory to save the uploaded files
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create the upload directory if it doesn't exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('upload.html')

def map_style(style, size):
    # Mapping logic based on style and size
    if style in ['CO129', 'CO078', 'CO079', 'HB2137', 'MK3514', 'MK3636', 'MK3558', 'MK3467', 'MK8558', 'MK3514KID', 'MK5178KID'] and size in ['1X', '2X', '3X', '4X']:
        if style == 'CO129':
            return 'CO129PL222'
        elif style == 'CO078':
            return 'CO078PL'
        elif style == 'CO079':
            return 'CO079PL'
        elif style == 'HB2137':
            return 'HB2137PL222'
        elif style == 'MK3514':
            return 'MK3514PL'
        elif style == 'MK3636':
            return 'MK3636Y'
        elif style == 'MK3558':
            return 'MK3558Y'
        elif style == 'MK3467':
            return 'MK3467Y'
        elif style == 'MK8558':
            return 'MK8558Y'
        elif style == 'MK3514KID':
            return 'MK3514KID'
        elif style == 'MK5178KID':
            return 'MK5178KID'

    # Return the original style if no conditions are met
    return style

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    
    file = request.files['file']

    if file.filename == '':
        return "No selected file"
    
    if file:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # Process the Excel file
        df = pd.read_excel(filepath)
        
        # Inspect column names to debug
        print(df.columns)
        
        # Safely split 'sku' and handle missing parts by padding/truncating
        sku_split = df['sku'].str.split('-', expand=True)
        
        # If there are less than 3 parts, pad with NaN. If more, truncate.
        df['Style'] = sku_split[0]
        df['Color'] = sku_split[1]
        df['Size'] = sku_split[2]

        # Define the second style mapping
        style_mapping = {
            'CB0536': 'CB0536',
            'CO078': 'CO078M',
            'CO078LEO': 'CO078Y/LEO',
            'CO079': 'CO079M',
            'CO129': 'CO129Y',
            'CO129LEO': 'CO129Y/LEO',
            'HB2122': 'HB2122',
            'HB2137': 'HB2137',
            'HB2137PL': 'HB2137PL222',
            'HB3134': 'HB3134',
            'HK8072': 'HK8072',
            'HK8246': 'HK8246',
            'HK8266': 'HK8266',
            'KC003': 'KC003',
            'KC009': 'KC009',
            'MK0179': 'MK0179',
            'MK3104': 'MK3104',
            'MK3279': 'MK3279',
            'MK3392': 'MK3392',
            'MK3466': 'MK3466',
            'MK3467': 'MK3467Y',
            'MK3506': 'MK3506',
            'MK3514': 'MK3514Y',
            'MK3515': 'MK3515',
            'MK3554': 'MK3554',
            'MK3558': 'MK3558Y',
            'MK3595': 'MK3595',
            'MK3637': 'MK3637Y',
            'MK3636': 'MK3636Y',
            'MK3659': 'MK3659',
            'MK3664': 'MK3664Y',
            'MK3664LEO': 'MK3664Y/LEO',
            'MK3673': 'MK3673',
            'MK3675': 'MK3675',
            'MK5178': 'MK5178',
            'MK5500': 'MK5500',
            'MK5502': 'MK5502',
            'MK8015': 'MK8015',
            'MK8080': 'MK8080',
            'MK8144': 'MK8144',
            'MK8213': 'MK8213',
            'MK8236': 'MK8236',
            'MK8558': 'MK8558Y',
            'MK5501': 'MK5501',
            'MK3664EMBO': 'MK3664EMBO',
            'MK32004CAT': 'MK32004CAT',
            'MK8281': 'MK8281',
            'MK3399': 'MK3399',
            'MK8143': 'MK8143',
            'MK8268': 'MK8268',
            'MK3349': 'MK3349',
        }    

        # Apply the second style mapping using the map_style function
        df['Style'] = df.apply(lambda row: map_style(row['Style'], row['Size']), axis=1)
        df['Style'] = df['Style'].replace(style_mapping)
        # Modify the 'Color' column based on a color mapping
        color_mapping = {
                'APL': 'APPLE',
                'APPLE': 'APPLE',
                'AQA': 'AQUA',
                'AQUA': 'AQUA',
                'BABYYELLOW': 'BABY YELLOW',
                'BBR': 'BLACKBERRY',
                'BER': 'BERRY',
                'BERRY': 'BERRY',
                'BGR': 'B.GREEN',
                'BKB': 'BLACKBERRY',
                'BLACK': 'BLACK',
                'BLACK/CORAL': 'BLACK/CORAL',
                'BLACK/CORK': 'BLACK/CORK',
                'BLACK/IVORY': 'BLACK/IVORY',
                'BLACK/MAUVE': 'BLACK/MAUVE',
                'BLACK/PINK': 'BLACK/PINK',
                'BLACK/RED': 'BLACK/RED',
                'BLACK/WHITE': 'BLACK/WHITE',
                'BLACKBERRY': 'BLACKBERRY',
                'BLACKBERRY/IVORY': 'BLACKBERRY/IVORY',
                'BLB': 'BLUEBERRY',
                'BLK':'BLACK',
                'BLK/IVR': 'BLACK/IVORY',
                'BLK/PNK': 'BLACK/PINK',
                'BLK/RED': 'BLACK/RED',
                'BLS': 'BLUSH',
                'BLS/GRY': 'BLUSH/GREY',
                'BLS/IVR': 'BLUSH/IVORY',
                'BLU': 'BLUE',
                'BLUE': 'BLUE',
                'BLUEBERRY': 'BLUEBERRY',
                'BLUEBERRY/LILAC': 'BLUEBERRY/LILAC',
                'BLUSH': 'BLUSH',
                'BLUSH/IVORY': 'BLUSH/IVORY',
                'BRG': 'B.GREEN',
                'BRICK': 'BRICK',
                'BRIGHTGREEN': 'B.GREEN',
                'BRK': 'BRICK',
                'BRONZE': 'BRONZE',
                'BRONZE/BLACK': 'BRONZE/BLACK',
                'BROWN': 'BROWN',
                'BRW': 'BROWN',
                'BRZ': 'BRONZE',
                'BRZ/BLK': 'BRONZE/BLACK',
                'BUR': 'BURGUNDY',
                'BURGUNDY': 'BURGUNDY',
                'BYL': 'BABY YELLOW',
                'CAM': 'CAMEL',
                'CAMEL': 'CAMEL',
                'CAMEL/BLACK': 'CAMEL/BLACK',
                'CAP': 'CAPRI',
                'CAPRI': 'CAPRI',
                'CCA': 'COCOA',
                'CHA': 'CHARCOAL',
                'CHARCOAL': 'CHARCOAL',
                'CLAY': 'CLAY',
                'CLY': 'CLAY',
                'COCOA': 'COCOA',
                'COF': 'COFFEE',
                'COFFEE': 'COFFEE',
                'COPPER': 'COPPER',
                'COR': 'CORAL',
                'CORAL': 'CORAL',
                'CORK': 'CORK',
                'CUS': 'CUSTARD',
                'D.ORANGE':'DUSTY ORANGE',
                'DCORAL': 'D.CORAL',
                'DCR': 'D.CORAL',
                'DCR': 'D.CORAL (CORAL)',
                'DOR': 'DUSTY ORANGE',
                'DOR/BLK': 'D.ORANGE/BLACK',
                'DUSTYCORAL': 'D.CORAL',
                'DUSTYORANGE': 'DUSTY ORANGE',
                'DUSTYORANGE/BLACK': 'D.ORANGE/BLACK',
                'FIESTA': 'FIESTA',
                'FOG': 'FOG',
                'FST': 'FIESTA',
                'FUCHSIA': 'FUCHSIA',
                'GOLD': 'GOLD',
                'GRAPE': 'GRAPE',
                'GRAY': 'GREY',
                'GRN': 'GREEN',
                'GREEN': 'GREEN',
                'GREY': 'GREY',
                'GREY': 'GRAY',
                'GREY/BLACK': 'GREY/BLACK',
                'GREY/IVORY': 'GREY/IVORY',
                'GREY/RED': 'GREY/RED',
                'GREY/WHITE': 'GREY/WHITE',
                'GRN': 'GREEN',
                'GRP': 'GRAPE',
                'GRY': 'GREY',
                'GRY': 'GRAY',
                'GRY': 'GREY',
                'GRY/IVY': 'GREY/IVORY',
                'H.CHARCOAL': 'HEATHER CHARCOAL',
                'H.CHARCOAL': 'H. CJARCOAL',
                'H.GREY': 'HEATHER GREY', 
                'HGR': 'HUNTER GREEN',
                'HGR': 'H.GREY',
                'HON': 'HONEY',
                'HON/BLK': 'HONEY/BLACK',
                'HON/PUP': 'HONEY/PURPLE',
                'HONEY': 'HONEY',
                'HONEY/IVORY': 'HONEY/IVORY',
                'HUNTERGREEN': 'HUNTER GREEN',
                'IBL': 'ICE BLUE',
                'ICEBLUE': 'ICE BLUE',
                'INK': 'INK',
                'IVORY': 'IVORY',
                'IVORY/BLACK': 'IVORY/BLACK',
                'IVORY/GRAY': 'IVORY/GREY',
                'IVORY/GREY': 'IVORY/GREY',
                'IVORY/RED': 'IVORY/RED',
                'IVORY/TAUPE': 'IVORY/TAUPE',
                'IVORY/TUAPE': 'IVORY/TAUPE',
                'IVR': 'IVORY',
                'IVR/BLK': 'IVORY/BLACK',
                'IVR/GRY': 'IVORY/GREY',
                'IVR/RED': 'IVORY/RED',
                'IVR/TPE': 'IVORY/TAUPE',
                'JAD': 'JADE',
                'JADE': 'JADE',
                'JDAE': 'JADE',
                'KELLYGREEN': 'KELLYGREEN',
                'KELLYGREEN': 'KELLYGREEN',
                'KELLYGREEN/IVORY': 'KELLYGREEN/IVORY',
                'KGR': 'KELLYGREEN',
                'KGR': 'KELLYGREEN',
                'L.ORANGE': 'L.ORANGE',
                'LAV': 'LAVENDER',
                'LAVENDER': 'LAVENDER',
                'LBL': 'L.BLUE',
                'LBLUE': 'L.BLUE',
                'LEMON': 'YELLOW',
                'LGR': 'L.GREY',
                'LIGHTBLUE': 'L.BLUE',
                'LIGHTGRAY': 'L.GREY',
                'LIGHTGREY': 'L.GREY',
                'LIGHTGREY/ORANGE': 'L.GREY/ORANGE',
                'LIGHTORANGE': 'L.ORANGE',
                'LIGHTPINK': 'L.PINK',
                'LIGHTPINK': 'LIGHT PINK',
                'LIL': 'LILAC',
                'LILAC': 'LILAC',
                'LOR': 'L.ORANGE',
                'LOR/IVR': 'L.ORANGE/IVORY',
                'LPINK': 'L.PINK',
                'LPK': 'L.PINK',
                'LPK': 'LIGHT PINK',
                'MAG': 'MAGENTA',
                'MAGENTA': 'MAGENTA',
                'Magenta': 'MAGENTA',
                'MAGENTA/BLACK': 'MAGENTA/BLACK',
                'MAR': 'MAROON',
                'MAUVE': 'MAUVE', 
                'MAUVE/BLACK': 'MAUVE/BLACK',
                'MAUVE/IVORY': 'MAUVE/IVORY',
                'MAUVEORCHID': 'MAUVE ORCHID',
                'MAV': 'MAUVE',
                'MAV/IVR': 'MAUVE/IVORY',
                'MCH': 'MOCHA',
                'MGT': 'MAGENTA',
                'MINT': 'MINT',
                'MOCHA': 'MOCHA',
                'MOR': 'MAUVE ORCHID',
                'MOS': 'MOSS',
                'MOS/IVR': 'MOSS/IVORY',
                'MOSS': 'MOSS',
                'MOSS/IVORY':'MOSS/IVORY',
                'MSH': 'MUSHROOM',
                'MUS': 'MUSTARD',
                'NAT': 'NATURAL',
                'NAV': 'NAVY',
                'NAV/OAT': 'NAVY/OATMEAL',
                'NAVY': 'NAVY',
                'NAVY/OATMEAL': 'NAVY/OATMEAL',
                'OAT': 'OATMEAEL',
                'OAT': 'OATMEAEL',
                'OAT/BLK': 'OATMEAL/BLACK',
                'OAT/PNK': 'OATMEAL/PINK',
                'OATMEAL': 'OATMEAEL',
                'OATMEAL': 'OATMEAEL',
                'OATMEAL/BLACK': 'OATMEAL/BLACK',
                'OATMEAL/GREY': 'OATMEAL/GREY',
                'OATMEAL/ORANGE': 'OATMEAL/ORANGE',
                'OATMEAL/PINK': 'OATMEAL/PINK',
                'OCH': 'ORCHID',
                'OLIVE': 'OLIVE',
                'OLIVE/BLACK': 'OLIVE/BLACK',
                'OLV': 'OLIVE',
                'OLV/BLK': 'OLIVE/BLACK',
                'OLV/HON': 'OLIVE/HONEY',
                'OPAL': 'OPAL',
                'OPL': 'OPAL',
                'ORANGE': 'ORANGE',
                'ORC': 'ORCHID',
                'ORCHID': 'ORCHID',
                'P.BEIGE/BLACK': 'P.BEIGE/BLACK',
                'PAPAYA': 'PAPAYA',
                'PBG': 'PEACH BEIGE',
                'PCH': 'PEACH',
                'PCK': 'PEACOCK',
                'PCK/HON': 'PEACOCK/HONEY',
                'PCK/IVR': 'PEACOCK/IVORY',
                'PEACH': 'PEACH',
                'Peach Beige': 'PEACH BEIGE',
                'PEACHBEIGE': 'PEACH BEIGE',
                'PEACHNECTAR': 'PEACH NECTOR',
                'PEACOCK': 'PEACOCK',
                'Peacock': 'PEACOCK',
                'PEACOCK/RED': 'PEACOCK/RED',
                'PINK': 'PINK',
                'PNK': 'PINK',
                'PNT': 'PEACH NECTOR',
                'PPA': 'PAPAYA',
                'PUP': 'PURPLE',
                'PURPLE': 'PURPLE',
                'RBL': 'ROYAL BLUE',
                'RBL': 'R/BLUE',
                'RED': 'RED',
                'RED/BLACK': 'RED/BLACK',
                'RED/BLK': 'RED/BLACK',
                'RED/IVORY': 'RED/IVORY',
                'RED/IVR': 'RED/IVORY',
                'REDPINK': 'RED PINK', 
                'ROYAL BLUE': 'ROYAL BLUE',
                'ROYALBLUE': 'ROYAL BLUE',
                'RPK': 'RED PINK',
                'RPK': 'ROSE PINK',
                'RST': 'RUST',
                'RUST': 'RUST',
                'SAG': 'SAGE',
                'SAGE': 'SAGE',
                'Sage': 'SAGE',
                'SAGE/BLACK': 'SAGE/BLACK',
                'SAL': 'SALMON',
                'SAND': 'SAND',
                'SBL': 'SKY BLUE',
                'SGR': 'SPRING GREEN',
                'SIL': 'SILVER',
                'SKYBLUE': 'SKY BLUE',
                'SND': 'SAND',
                'SPK': 'SWEET PINK',
                'SPRINGGREEN': 'SPRING GREEN',
                'SWEETPINK': 'SWEET PINK',
                'TAN': 'TAN',
                'TAN/BLACK':  'TAN/BLACK',
                'TAN/BLK': 'TAN/BLACK',
                'TAN/PNK': 'TAN/PINK',
                'TAUPE':  'TAUPE',
                'Taupe': 'TAUPE',
                'TAUPE/BLACK': 'TAUPE/BLACK',
                'TBL': 'TEAL BLUE',
                'TBL/IVR': 'TEAL BLUE/IVORY',
                'TEAL': 'TEAL',
                'Teal Blue': 'TEAL BLUE',
                'TEAL/BLACK': 'TEAL/BLACK',
                'TEALBLUE': 'TEAL BLUE',
                'TEALBLUE/WHITE': 'TEAL BLUE/WHITE',
                'TEL': 'TEAL',
                'TGR': 'TROPICAL GREEN',
                'TMT': 'TOMATO',
                'TMT/OAT': 'TOMATO/OATMEAL',
                'TOMATO': 'TOMATO',
                'TPE': 'TAUPE',
                'TQS': 'TURQUOISE',
                'TROPICALGREEN': 'TROPICAL GREEN',
                'TURQUOISE': 'TURQUOISE',
                'VIL': 'VIOLA',
                'VIO': 'VIOLA',
                'VIOLA': 'VIOLA',
                'VIOLET': 'VIOLET',
                'VLT': 'VIOLET',
                'WHITE': 'WHITE',
                'White': 'WHITE',
                'WHT':'WHITE',
                'YEL': 'YELLOW',
                'YELLOW': 'YELLOW',
                'MSR': 'MUSHROOM',
                'NVY': 'NAVY',
                'KGN': 'KELLYGREEN'
            }
        df['Color'] = df['Color'].replace(color_mapping)

        # Convert sizes 'S/M' and 'M/L' to 'SM' and 'ML'
        df['Size'] = df['Size'].replace({'S/M': 'SM', 'M/L': 'ML'})

        # Prepare the Excel file to be downloaded
        output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'output.xlsx')
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "FirstLayout"

        # Write the headers
        headers = ['Style', 'Color', 'Size', 'Quantity']
        ws1.append(headers)

        # Apply some formatting
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        for col in range(1, 5):
            cell = ws1.cell(row=1, column=col)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

        # Write the data
        for index, row in df.iterrows():
            ws1.append([row['Style'], row['Color'], row['Size'], row['quantity-purchased']])

        # Save the workbook
        wb.save(output_filepath)

        return send_file(output_filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
