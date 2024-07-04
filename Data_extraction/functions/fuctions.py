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
import subprocess


dotenv.load_dotenv()
openai_api_key = str(os.getenv("OPENAI_API_KEY"))
client = openai.Client(api_key=openai_api_key)
MODEL = "gpt-4o"

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

def extract_images_from_pdf(pdf_path):
    try:
        document = fitz.open(pdf_path)
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        images_dir = os.path.join('image_files', pdf_name)
        
        if not os.path.exists(images_dir):
            os.makedirs(images_dir)
        
        image_count = 0
        
        for page_index in range(len(document)):
            page = document[page_index]
            image_list = page.get_images(full=True)
            
            
            for image_index, img in enumerate(image_list):
                if image_count >= 3:
                    break
                
                xref = img[0]
                base_image = document.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                width = base_image["width"]
                height = base_image["height"]
                
                if not image_bytes:
                    continue 
                if width <= 200 or height <= 200:
                    continue   
                image_count += 1
                image_path = os.path.join(images_dir, f'image{image_count}.{image_ext}')
                with open(image_path, "wb") as image_file:
                    image_file.write(image_bytes)

                if image_ext.lower() in ['jpx', 'jpeg2000', 'jp2']:
                    image_path_jpg = os.path.join(images_dir, f'image{image_count}.jpg')
                    try:
                        subprocess.run(['magick', 'convert', image_path, image_path_jpg])
                        print(f"Converted image to jpg: {image_path_jpg}")
                        os.remove(image_path)
                        image_path = image_path_jpg
                    except Exception as e:
                        print(f"Error converting {image_path} to JPEG: {e}")
            
            if image_count >= 3:
                break
    
    except Exception as e:
        print(f"Error extracting images from PDF: {e}")

 
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

def generate_image_links(api_key,pdf_path):
    filename = os.path.basename(pdf_path)
    filename = os.path.splitext(filename)[0]
    
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
    image_files = [os.path.join('image_files', pdf_name, image_file) for image_file in os.listdir(os.path.join('image_files', pdf_name))]
    print(len(image_files))
    image_links = []

    for image_file in image_files:
        if os.path.exists(image_file):
            try:
                link = upload_image_to_freeimage(image_file, api_key)
                image_links.append(link)
            except Exception as e:
                print(f"Error uploading {image_file}: {e}")
        else:
            print(f"{image_file} does not exist.")
    
    return image_links

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)
        text = page.get_text()
        doc.close()
    except Exception as e:
        print(f"Error extracting text from PDF: {e}")
    return text

 
def extract_specifications(specification):
    try:
        specification_text = "\n".join(specification)
        messages = [
            # we can change the prompt in the following line
            {"role": "system", "content": "Extract specifications in HTML table format carefully as the cell maybe out of order. no th headings just td tags in table and all td tags should have some content Inner text.  If 2 first td's in a tr have same name then make only one tr and put the second tds in the same place with br tag. If found multiple SKU's or products, keep them in same table with multiple tds/column each model seperated by row/<tr> each with their own specifications Return in <table>. If there is no table found in the input, return 'None' as the output. "},
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
    
def extract_features_list(raw_text):
    try:
        messages = [
            # we can change the prompt in the following line
            {"role": "system", "content": "Extract features in HTML list format. Use 'li' tags for each feature inside the ul Features. In the raw text, the features are listed in bullet points. Extract the features and present them in an HTML list. If there are multiple products or SKUs, extract the features for the given product only. If the features are not in bullet points, generate them as bullet points. Not more than 10 features. Return each feature in an <li> tag inside a <ul> tag."},
            {"role": "user", "content": raw_text}
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
        messages = [
            # we can change the prompt in the following line
            {"role": "system", "content": "Extract the description of the product. The description should be in a single paragraph. and not more than 4 lines see the text and extract description."},
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
    
def extract_name_from_ai(raw_text):
    try:
        messages = [
            {"role": "system", "content": "Extract the name of the product. If you cannot find the name from the text provided, generate it with Company Name + Model Number + Product Type(Any Machine or something else)  + Voltage(If found)"},
            {"role": "user", "content": raw_text}
        ]
        response = client.chat.completions.create(
            model=MODEL,
            messages=messages,
            temperature=0.0  
        )
        name = response.choices[0].message.content
        
        return name  
    except Exception as e:
        print(f"Error extracting name: {e}")
        return "Error extracting name."
    
def extract_specification_from_table(pdf_path):
    try:
        specifications = []
        doc = fitz.open(pdf_path)

        def extract_tables_from_page(page):
            tables = page.find_tables()
            if tables:
                for i, tab in enumerate(tables):
                    if i == 0:
                        continue
                    table_content = tab.extract()
                    for table_row in table_content:
                        for cell in table_row:
                            if cell and cell.strip() != '':
                                specifications.append(cell)
            else:
                print("No tables found on page.")

        first_page = doc.load_page(0)
        extract_tables_from_page(first_page)

        if not specifications and len(doc) > 1:
            second_page = doc.load_page(1)
            extract_tables_from_page(second_page)

        doc.close()
        
        if not specifications:
            print("No tables found in the first two pages.")
            
    except Exception as e:
        print(f"Error extracting table: {e}")

    return specifications

def update_excel_row(file_path, row_index, image_links, specifications,name,features,description):
    try:
        df = pd.read_excel(file_path)
        df.at[row_index, 'Name'] = name
        df.at[row_index, 'Features'] = features
        df.at[row_index, 'Description'] = description
        df.at[row_index, 'Specifications'] = specifications
        for i in range(3):
            if i < len(image_links):
                df.at[row_index, f'Image {i+1}'] = image_links[i]
            else:
                df.at[row_index, f'Image {i+1}'] = ""
        df.to_excel(file_path, index=False)
        print(f"Updated Excel row {row_index} successfully.")
    except Exception as e:
        print(f"Error updating Excel row {row_index}: {e}")

def extract_name_method_1(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)
        page_height = page.rect.height
        top_margin = page_height * 0.2
        bottom_margin = page_height * 0.3
        screenshot_rect = fitz.Rect(0, top_margin, page.rect.width, bottom_margin)
        text = page.get_text("text", clip=screenshot_rect)
        doc.close()
        text = text.split("\n")
        name = text[0]
    except Exception as e:
        print(f"Error extracting name (method 1): {e}")
        name = " "
    return name

def extract_name_method_2(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)
        page_height = page.rect.height
        top_margin = page_height * 0.1
        bottom_margin = page_height * 0.2
        screenshot_rect = fitz.Rect(0, top_margin, page.rect.width, bottom_margin)
        text = page.get_text("text", clip=screenshot_rect)
        doc.close()
        text = text.split("\n")
        name = text[0]
    except Exception as e:
        print(f"Error extracting name (method 2): {e}")
        name = " "
    return name

def extract_name(pdf_path,raw_text):
    name = extract_name_method_1(pdf_path)
    if len(name.split()) < 6 or len(name.split()) > 12:
        name = extract_name_method_2(pdf_path)
    if name == ' ' :
        name = extract_name_from_ai(raw_text)
    return name

if __name__ == "__main__":
    file_path = 'C:/Work/upwork/price-scraper/Data_extraction/excel_file/USR.xlsx'
    images_api_key = str(os.getenv("IMAGES_API_KEY"))
    df = pd.read_excel(file_path)
    images_api_key = str(os.getenv("IMAGES_API_KEY"))
    filtered_df = df[df['Specsheet'].notna()]
    start_row = 241
    end_row = 328
    for idx, row in filtered_df.iterrows():
        pdf_url = row['Specsheet']
        pdf_path = download_pdf(pdf_url)
        raw_text = extract_text_from_pdf(pdf_path)
        if not raw_text:
            print(f"Skipping row {idx} due to failed text extraction.")
            continue
        name = extract_name(pdf_path, raw_text)
        features = extract_features_list(raw_text)
        specification = extract_specification_from_table(pdf_path)
        description = extract_description(raw_text)
        specifications = extract_specifications(specification)
        extract_images_from_pdf(pdf_path)
        image_links = generate_image_links(images_api_key, pdf_path)
        update_excel_row(file_path, idx, image_links, specifications,name,features,description)



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
   


