import requests
import fitz  
import os
import os
from pdf2image import convert_from_path
import base64
## Set the API key
import os
import openai
from openai import OpenAI
import dotenv
import pandas as pd
from PIL import Image, ImageFile
import io
from spire.pdf.common import *
from spire.pdf import *

dotenv.load_dotenv()
openai_api_key = str(os.getenv("OPENAI_API_KEY"))
print(openai_api_key)
client = openai.Client(api_key=openai_api_key)
MODEL = "GPT-4o"

def download_pdf(url):
    response = requests.get(url)
    directory = "pdf_files"
    if not os.path.exists(directory):
        os.makedirs(directory)
    filename = os.path.basename(url)
    pdf_path = os.path.join(directory, filename)

    with open(pdf_path, 'wb') as f:
        f.write(response.content)
    
    return pdf_path

# download_pdf(urls)

# for the first five pdfs we need to extract the text from the first five pdfs

ImageFile.LOAD_TRUNCATED_IMAGES = True

def capture_screenshots(pdf_path):
    try:
        pdf_file = fitz.open(pdf_path)
        if not os.path.exists("image_files"):
            os.makedirs("image_files")

        for page_number in range(len(pdf_file)):
            page = pdf_file.load_page(page_number)
            page_width = page.rect.width
            page_height = page.rect.height
            
            if page_number == 0:
        
                reduced_height = page_height * 0.6
                top_margin = (page_height - reduced_height) / 2
                bottom_margin = page_height - top_margin
            
                screenshot_rect = fitz.Rect(0, top_margin, page_width // 2, bottom_margin)
                image_path = f"image_files/page_{page_number+1}_image_1.jpg"
                
        
                pixmap = page.get_pixmap()
                img = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
                
                
                img_cropped = img.crop((screenshot_rect.x0, screenshot_rect.y0, screenshot_rect.x1, screenshot_rect.y1))
                img_cropped.save(image_path)
                print(f"Image saved: {image_path}")

            elif page_number == 1:
                # Reduce page height by 20% from top and bottom
                reduced_height = page_height * 0.8
                top_margin = (page_height - reduced_height) / 2
                bottom_margin = page_height - top_margin
                
                # Right half of the second page, divided into two halves
                right_half_width = page_width // 2
                screenshot_rect_1 = fitz.Rect(right_half_width, top_margin, page_width, top_margin + reduced_height / 2)
                screenshot_rect_2 = fitz.Rect(right_half_width, top_margin + reduced_height / 2, page_width, bottom_margin)
                
                # Render page to image
                pixmap = page.get_pixmap()
                img = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
                
                # Crop and save top half of the right side
                img_cropped_1 = img.crop((screenshot_rect_1.x0, screenshot_rect_1.y0, screenshot_rect_1.x1, screenshot_rect_1.y1))
                image_path_1 = f"image_files/page_{page_number+1}_image_2.jpg"
                img_cropped_1.save(image_path_1)
                print(f"Image saved: {image_path_1}")

                # Crop and save bottom half of the right side
                img_cropped_2 = img.crop((screenshot_rect_2.x0, screenshot_rect_2.y0, screenshot_rect_2.x1, screenshot_rect_2.y1))
                image_path_2 = f"image_files/page_{page_number+1}_image_3.jpg"
                img_cropped_2.save(image_path_2)
                print(f"Image saved: {image_path_2}")

            else:
                break  # Only handle the first two pages as per your requirement

        pdf_file.close()
    
    except Exception as e:
        print(f"Error capturing screenshots from PDF: {e}")
 
def upload_image_to_freeimage(image_path, api_key):
    """Uploads an image to FreeImage.host and returns the direct image link."""
    url = "https://freeimage.host/api/1/upload"
    with open(image_path, 'rb') as image_file:
        files = {
            'source': image_file,
        }
        data = {
            'key': api_key,
            'action': 'upload',
            'type': 'file',
        }
        response = requests.post(url, files=files, data=data)
        if response.status_code == 200:
            response_data = response.json()
            if response_data['status_code'] == 200:
                return response_data['image']['url']
            else:
                raise Exception(f"Failed to upload image: {response_data['status_code']} {response_data['status_txt']}")
        else:
            raise Exception(f"Failed to upload image: {response.status_code} {response.text}")

def generate_image_links(api_key):
    image_files = ["image_files/page_1_image_1.jpg", "image_files/page_2_image_2.jpg", "image_files/page_2_image_3.jpg"]
    image_links = {}

    for image_file in image_files:
        if os.path.exists(image_file):
            try:
                link = upload_image_to_freeimage(image_file, api_key)
                image_links[image_file] = link
                print(f"Uploaded {image_file} to {link}")
            except Exception as e:
                print(f"Error uploading {image_file}: {e}")
        else:
            print(f"{image_file} does not exist.")
    
    return image_links

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        doc = fitz.open(pdf_path)
        # Extract text from each page
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text += page.get_text()
        doc.close()
    except Exception as e:
        print(f"Error extracting text from PDF: {e}")
    return text

def process_text(text, code):
    lines = text.split('\n')
    name = None
    features = []
    specification_lines = []
    features_started = False
    specification_started = False
    electrical_found = False
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if "FEATURES" in line:
            features_started = True
            continue
        if features_started:
            if line.startswith("•"):
                features.append(line.strip("• ").strip())
            else:
                features_started = False
            
        # Extract specification section
        if "CONSTRUCTION" in line:
            specification_started = True
        if specification_started:
            if "ELECTRICAL" in line:
                electrical_found = True
            specification_lines.append(line.strip())
            if electrical_found and "COOKING" in line:
                break
            if electrical_found and not "COOKING" in line and "USR Brands" in line:
                specification_lines.pop()
                break

    specification_text = '\n'.join(specification_lines).strip()

    for line in lines:
        if line.startswith(code) or code in line:
            name = line.strip()
            break

    return name, features,specification_text



# I want to parse the specification to gpt-4o model to extract the specifications in html table
def pdf_to_image(pdf_path, output_dir="image_files", dpi=300, page_numbers=None):
    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        if page_numbers is None:
            page_numbers = [0]  # Default to convert only the first page if not specified
        
        image_paths = []
        images = convert_from_path(pdf_path, dpi=dpi, first_page=min(page_numbers), last_page=max(page_numbers))

        for i, image in enumerate(images):
            image_path = os.path.join(output_dir, f"{os.path.basename(pdf_path).replace('.pdf', f'_{i+1}.png')}")
            image.save(image_path, 'PNG')
            print(f"Saved image: {image_path}")
            image_paths.append(image_path)
        
        return image_paths

    except Exception as e:
        print(f"Error converting PDF to image: {e}")
        return []
def extract_specifications(specification):
    try:
        print(f"Extracting specifications: {specification}")
        specification_text = specification.replace("\n", " ").strip()    
        messages = [
            {"role": "system", "content": "Extract specifications in HTML table format. no th headings just td tags in table"},
            {"role": "user", "content": specification_text}
        ]
        response = client.chat.completions.create(
            model=MODEL,
            messages=messages,
            temperature=0.0  
        )
        html_table = response.choices[0].message.content
        
        return html_table  
    except Exception as e:
        print(f"Error extracting specifications: {e}")
        return "Error extracting specifications."
def extract_features_list(features):
    try:
        print(f"Extracting features: {features}")
        features_text = "\n".join([f"- {feature}" for feature in features])
        messages = [
            {"role": "system", "content": "Extract features in HTML list format. Use 'li' tags for each feature inside the ul Features."},
            {"role": "user", "content": features_text}
        ]
        response = client.chat.completions.create(
            model=MODEL,
            messages=messages,
            temperature=0.0  
        )
        html_list = response.choices[0].message.content
        
        return html_list  
    except Exception as e:
        print(f"Error extracting features: {e}")
        return "Error extracting features."

def extract_description(raw_text):
    try:
        print(f"Extracting description: {raw_text}")
        messages = [
            {"role": "system", "content": "Extract the description of the product. The description should be in a single paragraph. and not more than 5 lines see the text and extract description."},
            {"role": "user", "content": raw_text}
        ]
        response = client.chat.completions.create(
            model=MODEL,
            messages=messages,
            temperature=0.0  
        )
        description = response.choices[0].message.content
        
        return description  
    except Exception as e:
        print(f"Error extracting description: {e}")
        return "Error extracting description."

def update_excel_row(file_path, row_index, name, description, features, specifications):
    try:
        df = pd.read_excel(file_path)
        df.at[row_index, 'Name'] = name
        df.at[row_index, 'Description'] = description
        df.at[row_index, 'Features'] = features
        df.at[row_index, 'Specifications'] = specifications
        df.to_excel(file_path, index=False)
        print(f"Updated Excel row {row_index} successfully.")
    except Exception as e:
        print(f"Error updating Excel row {row_index}: {e}")

if __name__ == "__main__":
    file_path = 'C:/Work/upwork/price-scraper/Data_extraction/excel_file/USR.xlsx'
    df = pd.read_excel(file_path)
    # pdf_path = 'pdf_files/bd250-specsheet.pdf'
    # capture_screenshots(pdf_path)
    images_api_key = str(os.getenv("IMAGES_API_KEY"))
    print(images_api_key)
    # image_links = generate_image_links(images_api_key)
    # print("Image Links:", image_links)
    # filtered_df = df[df['Specsheet'].notna()]
    # max_rows_to_process = 5  

    # for idx, row in filtered_df.head(max_rows_to_process).iterrows():
    #     pdf_url = row['Specsheet']
    #     pdf_path = download_pdf(pdf_url)
    #     code = row['SKU']

    #     raw_text = extract_text_from_pdf(pdf_path)
    #     name, features, specification = process_text(raw_text, code)
    #     description = extract_description(raw_text)
    #     specifications = extract_specifications(specification)
    #     features_html_list = extract_features_list(features)

    #     update_excel_row(file_path, idx, name, description, features_html_list, specifications)



# def download_pdf(pdf_url):
#     response = requests.get(pdf_url)
#     directory = "pdf_files"
#     if not os.path.exists(directory):
#         os.makedirs(directory)
#     pdf_path = os.path.join(directory, "downloaded.pdf")
#     with open(pdf_path, 'wb') as f:
#         f.write(response.content)
#     return pdf_path

# def extract_text_from_pdf(pdf_path):
#     """
#     Extracts text from a specific page of a PDF file and returns it as a string.
    
#     :param pdf_path: Path to the PDF file
#     :param page_number: The page number to extract text from (1-based index)
#     :return: Extracted text as a string
#     """
#     page_number=1
#     document = fitz.open(pdf_path)
    
#     if page_number < 1 or page_number > len(document):
#         return f"Page number {page_number} is out of range. The document has {len(document)} pages."
    
#     page = document.load_page(page_number - 1)  # Convert to 0-based index
#     text = page.get_text()
    
#     document.close()
    
#     return text

# output_folder='images'

# def pdf_to_image(pdf_path, dpi=300):
#     # Ensure the output folder exists
#     if not os.path.exists(output_folder):
#         os.makedirs(output_folder)
    
#     # Convert all pages of the PDF
#     images = convert_from_path(pdf_path, dpi=dpi)
   
#     # Save each page as a separate image
#     if images:
#         image_path = os.path.join(output_folder, 'page_1.png')
#         images[0].save(image_path, 'PNG')
#         print(f'Saved: {image_path}')
#     else:
#         print('No images were created.')


# # Open the image file and encode it as a base64 string
# def encode_image(image_path):
#     with open(image_path, "rb") as image_file:
#         return base64.b64encode(image_file.read()).decode("utf-8")


# def extract_specifications(image_base64,code):
#     response = client.chat.completions.create(
#     model=MODEL,
#     messages=[
#         {"role": "system", "content": "You are a helpful assistant that responds in text   Extract the specifications only and present them in an HTML table. Ensure the output is in HTML format please do not add headings in specifications , not as a string. Provide only the specifications. data "},
        
# #         {"role": "system", "content": "Write a detailed description for a product. The description should be saved in the 'description' key with the description as the value. Return the result in JSON format."},
#          {"role": "system", "content": "do not include Features details in this "},
#          {"role": "system", "content": f"if there are multiples products sku then extract details of given product {code}  code  "},
#         {"role": "user", "content": [


#             {"type": "image_url", "image_url": {
#                 "url": f"data:image/png;base64,{image_base64}"}
#             }
#         ]}
#     ],
#     temperature=0.0
# )
#     Specification=(response.choices[0].message.content)
#     return Specification


# def extract_features_to_html(image_base64,code):
#     response = client.chat.completions.create(
#     model=MODEL,
#     messages=[
#         {"role": "system", "content": "You are a helpful assistant that responds in text   Extract the Features only and present them in an HTML table. Ensure the output is in HTML list (li), not as a string. Provide only the Features.  "},
#          {"role": "system", "content": f"if there are multiples products sku then extract details of given product {code}  code  "},
#         {"role": "user", "content": [


#             {"type": "image_url", "image_url": {
#                 "url": f"data:image/png;base64,{image_base64}"}
#             }
#         ]}
#     ],
#     temperature=0.0
# )
    
#     features =(response.choices[0].message.content)
#     return features




# def extract_product_details(image_base64,code):
#     response = client.chat.completions.create(
#     model=MODEL,
#     messages=[
#         {"role": "system", "content": "You are a helpful assistant that responds in text   Extract the full product name mandatory,  Certifications ,Warranty in dictionary  form it should be in pure josn form not make string not include like```json dont add incorrect details   "},
#         {"role": "system", "content": "while extracting warrenty make give values like example Warranty: 2 years parts & 1 year labor exclusive of lights, gaskets and glass thenn value should be  '1:Labor, 2:Parts'"},
#         {"role": "system", "content": "while extracting warrenty make give values like example Certifications ,please extract it froms logs like this ETL-Sanitation, cETLus,'if you find the names in logo then do not add  from text' make sure certification should return like ETLSAN ,CETLUS"},
#          {"role": "system", "content": f"if there are multiples products sku then extract details of given product {code}  code  "},
#         {"role": "system", "content": "The product name should match exactly as it appears in the image. Please provide the full and accurate product name."},
        
#         {"role": "user", "content": [


#             {"type": "image_url", "image_url": {
#                 "url": f"data:image/png;base64,{image_base64}"}
#             }
#         ]}
#     ],
#     temperature=0.0
# )
#     data3=(response.choices[0].message.content)
#     return data3
   


