import json
import xml.etree.ElementTree as ET
import os
import sys
import zipfile

# Author - Daniel Agafonov (https://github.com/AlphaHasher)
# Author of lines 89-103, 164-173 - Martijn Lentink (https://github.com/martijnlentink)

# THINGS TO DO AND UPDATE BEFORE USING THIS SCRIPT:
# 1. Download your desired translation from here: https://github.com/martijnlentink/propresenter-custom-bibles . This will give you an output folder which will have the USX file you need to target.
# 2. Update the file_path variable to the path of your USX file that you want to target (for example: the PSA.usx file in the output folder)
# 3. Update the output_folder variable to the path of the output folder that was created when you downloaded the translation from the link above
# 4. Update the translation_name variable to the name of the translation you are adjusting (for example: ESV RUS Adjusted)
# 5. Update the metadata and rvmetadata files to include the new translation name (this is where the translation name will be displayed in ProPresenter)

# KNOWN ISSUES:
# 1. Psalm chapter 9 will be broken because of the way the chapter numbers are adjusted. This is fixed by manually removing the second chapter 9 in the USX file.

# Load the USX file
file_path = "path/to/your/file.usx"
output_folder = "path/to/your/output/folder" # created after you download the translation from the link above
translation_name = "NAME RUS Adjusted"

tree = ET.parse(file_path)
root = tree.getroot()

# The verse_map dictionary
verse_map = {
    'PSA': {
        '1:1-2:12': 'X:X',
        '3:1-9:20': 'X:+1', # Issue occurs here where there are now two chapters labeled 9
        '10:1': '-1:+21',
        '10:2-18': '-1:+21',
        '11:1-7': '-1:X',
        '12:1-13:4': '-1:+1',
        '13:5': '-1:+1',
        '13:6': '-1:+1',
        '14:1-17:15': '-1:X',
        '18:1-22:31': '-1:+1',
        '23:1-29:11': '-1:X',
        '30:1-31:24': '-1:+1',
        '32:1-33:22': '-1:X',
        '34:1-22': '-1:+1',
        '35:1-28': '-1:X',
        '36:1-12': '-1:+1',
        '37:1-40': '-1:X',
        '38:1-42:11': '-1:+1',
        '43:1-5': '-1:X',
        '44:1-49:20': '-1:+1',
        '50:1-23': '-1:X',
        '51:1-52:9': '-1:+2',
        '53:1-6': '-1:+1',
        '54:1-7': '-1:+2',
        '55:1-59:17': '-1:+1',
        '60:1-12': '-1:+2',
        '61:1-65:13': '-1:+1',
        '66:1-20': '-1:X',
        '67:1-70:5': '-1:+1',
        '71:1-74:23': '-1:X',
        '75:1-77:20': '-1:+1',
        '78:1-79:13': '-1:X',
        '80:1-81:16': '-1:+1',
        '82:1-8': '-1:X',
        '83:1-85:13': '-1:+1',
        '86:1-17': '-1:X',
        '87:1': '-1:+1',
        '87:2': '-1:X',
        '87:3-7': '-1:X',
        '88:1-90:4': '-1:+1',
        '90:5': '-1:+1',
        '90:6': '-1:X',
        '90:7-91:16': '-1:X',
        '92:1-15': '-1:+1',
        '93:1-101:8': '-1:X',
        '102:1-28': '-1:+1',
        '103:1-107:43': '-1:X',
        '108:1-13': '-1:+1',
        '109:1-114:8': '-1:X',
        '115:1-18': '-2:+8',
        '116:1-9': '-2:X',
        '116:10-19': '-1:-9',
        '117:1-147:11': '-1:X',
        '147:12-20': 'X:-11',
        '148:1-150:6': 'X:X',
    }
}

def move_rvbible_propresenter_folder(rvbible_loc):
    system_str = platform.system()
    _, filename = os.path.split(rvbible_loc)

    if system_str == 'Windows':
        program_data = os.getenv('PROGRAMDATA')
        propresenter_bible_location = os.path.join(program_data, 'RenewedVision\ProPresenter\Bibles\sideload')
        os.makedirs(propresenter_bible_location, exist_ok=True)
    elif system_str == 'Darwin':
        propresenter_bible_location = '/Library/Application Support/RenewedVision/RVBibles/v2/'
    else:
        raise Exception("Unable to determine operating system, please copy the bible manually")

    new_file_loc = os.path.join(propresenter_bible_location, filename)
    shutil.copyfile(rvbible_loc, new_file_loc)

def apply_adjustment(chap_no, verse_no, adjustment):
    chap_adj, verse_adj = adjustment.split(":")

    if chap_adj == "X":
        new_chap_no = chap_no
    else:
        new_chap_no = str(int(chap_no) + int(chap_adj))

    if verse_adj == "X":
        new_verse_no = verse_no
    else:
        new_verse_no = str(int(verse_no) + int(verse_adj))

    return new_chap_no, new_verse_no

chapter_map = {}

for elem in root.iter():
    if elem.tag == "book":
        book = elem.text
        book_c = elem.get('code')

    if elem.tag == "chapter":
        chap_no = elem.get('number')
        chapter_map[chap_no] = elem  # Store chapter elements for later adjustment

    if elem.tag == "verse":
        verse_no = elem.get('number')
        ref = f"{chap_no}:{verse_no}"

        for ref_range, adjustment in verse_map.get(book_c, {}).items():

            if "-" not in ref_range:
                start_ref = end_ref = ref_range
            else:
                start_ref, end_ref = ref_range.split("-")
                start_chap, start_verse = map(int, start_ref.split(":"))

                if ":" not in end_ref:
                    end_ref = f"{start_chap}:{end_ref}"

                end_chap, end_verse = map(int, end_ref.split(":"))

            if (start_chap < int(chap_no) < end_chap) or (start_chap == int(chap_no) and start_verse <= int(verse_no)) or (end_chap == int(chap_no) and int(verse_no) <= end_verse):
                new_chap_no, new_verse_no = apply_adjustment(chap_no, verse_no, adjustment)
                elem.set('number', new_verse_no)
                chapter_map[chap_no] = new_chap_no  # Store the new chapter number
                break

        print(f"Psalm {chap_no}:{verse_no} adjusted to {new_chap_no}:{new_verse_no}")

# Update the chapter numbers in the XML structure
for old_chap_no, new_chap_no in chapter_map.items():
    for chap_elem in root.findall(f".//chapter[@number='{old_chap_no}']"):
        chap_elem.set('number', new_chap_no)
        print(f"Chapter {old_chap_no} updated to {new_chap_no}")

# Save the modified XML back to the file
tree.write(file_path, encoding='utf-8', xml_declaration=False)

zip_location = os.path.join(output_folder, f"../{translation_name}.rvbible")
shutil.make_archive(zip_location, 'zip', output_folder)
rvbible_location = shutil.move(zip_location + ".zip", zip_location)

print("Moving bible to ProPresenter directory")

move_rvbible_propresenter_folder(rvbible_location)

print("Done! Please restart ProPresenter and check if the bible is correctly installed.")
input("Press enter to close...")