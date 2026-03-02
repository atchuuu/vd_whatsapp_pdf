import io
import os
import json
import requests
import time
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import fitz
from PIL import Image, ImageChops
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from datetime import datetime
import boto3

GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON")

SHEET_ID = os.environ.get("SHEET_ID")

R2_ACCOUNT_ID = os.environ.get("R2_ACCOUNT_ID")
R2_ACCESS_KEY = os.environ.get("R2_ACCESS_KEY")
R2_SECRET_KEY = os.environ.get("R2_SECRET_KEY")
R2_BUCKET_NAME = os.environ.get("R2_BUCKET_NAME")
R2_PUBLIC_BASE = os.environ.get("R2_PUBLIC_BASE")

AISENSY_API_KEY = os.environ.get("AISENSY_API_KEY")
CAMPAIGN_NAME = os.environ.get("CAMPAIGN_NAME")

DESTINATIONS = [d.strip() for d in os.getenv("DESTINATIONS", "").split(",") if d.strip()]

TODAY = datetime.now().strftime("%d %B %Y")
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
FILE_NAME = f"VD_Report_{TIMESTAMP}.pdf"
SECTIONS =  [
    ("VD Report", "B8:X14", "Hello Team, Overall Leader level sales view summary along with their Overall targets."),
    ("VD Report", "B27:U33", "#Leader level sales view summary along with their day targets."),
    ("VD Top Batch Day View", "A5:F20", "#Top Batch Day 0 Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "L6:Q20", "#Top Batch YTD Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "A21:F37", "#Top Batch Day 1 Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "L23:Q37", "#Top Batch YTD Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "A38:F54", "#Top Batch Day 2 Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "L40:Q54", "#Top Batch YTD Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "A55:F71", "#Top Batch Day 3 Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "L57:Q71", "#Top Batch YTD Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "A72:F88", "#Top Batch Day 4 Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "L74:Q88", "#Top Batch YTD Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "A89:F105", "#Top Batch Day 5 Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "L91:Q105", "#Top Batch YTD Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "A106:F122", "#Top Batch Day 6 Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "L123:Q122", "#Top Batch YTD Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "A123:F139", "#Top Batch Day 7 Sales View with 2-year comparison and YoY growth."),
    ("VD Top Batch Day View", "L125:Q139", "#Top Batch YTD Sales View with 2-year comparison and YoY growth."),
]
print("✅ Environment Variables Loaded")

def get_google_creds():
    if not GOOGLE_CREDENTIALS_JSON:
        raise Exception("GOOGLE_CREDENTIALS_JSON environment variable is missing")
    info = json.loads(GOOGLE_CREDENTIALS_JSON)
    creds = Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    creds.refresh(Request())
    return creds

def get_sheet_gid(creds, sheet_name):
    service = build("sheets", "v4", credentials=creds)
    meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    for sheet in meta["sheets"]:
        if sheet["properties"]["title"] == sheet_name:
            return str(sheet["properties"]["sheetId"])
    raise Exception(f"Sheet {sheet_name} not found")

def trim_white_space(pil_img):
    bg = Image.new(pil_img.mode, pil_img.size, (255, 255, 255))
    diff = ImageChops.difference(pil_img, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        padding = 15
        b_x0 = max(0, bbox[0] - padding)
        b_y0 = max(0, bbox[1] - padding)
        b_x1 = min(pil_img.size[0], bbox[2] + padding)
        b_y1 = min(pil_img.size[1], bbox[3] + padding)
        return pil_img.crop((b_x0, b_y0, b_x1, b_y1))
    return pil_img

def export_range_image(creds, sheet_name, range_name):
    sheet_gid = get_sheet_gid(creds, sheet_name)

    export_url = (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export"
        f"?format=pdf&gid={sheet_gid}&range={range_name}&size=A2&portrait=true&fitw=true"
        f"&scale=2&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false"
        f"&fzr=false&top_margin=0&bottom_margin=0&left_margin=0&right_margin=0"
    )

    max_retries = 5
    for attempt in range(max_retries):
        response = requests.get(export_url, headers={"Authorization": f"Bearer {creds.token}"}, timeout=120)
        if response.status_code == 429:
            delay = 2 ** attempt * 2
            print(f"⚠️ Rate limited (429) for {range_name}. Retrying in {delay} seconds...")
            time.sleep(delay)
            continue
        response.raise_for_status()
        break
    else:
        response.raise_for_status()

    if not response.content.startswith(b"%PDF"):
        raise Exception("Invalid PDF returned from Google")

    doc = fitz.open(stream=response.content, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=450)
    img_bytes = pix.tobytes("png")
    
    pil_img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    cropped_img = trim_white_space(pil_img)
    
    cropped_bytes = io.BytesIO()
    cropped_img.save(cropped_bytes, format="PNG")
    cropped_bytes.seek(0)
    
    return ImageReader(cropped_bytes), cropped_img.width, cropped_img.height

def generate_dynamic_single_page_clean():
    creds = get_google_creds()



    images_data = []
    total_h = 0
    PAGE_WIDTH = 1800
    MARGIN = 70 
    USABLE_WIDTH = PAGE_WIDTH - (MARGIN * 2)

    print("📄 Capturing regions from Google Sheets...")
    for sheet_name, range_name, description in SECTIONS:
        print(f"   -> {sheet_name} ({range_name})")
        time.sleep(2)  # Delay to prevent rate limiting
        img_reader, w, h = export_range_image(creds, sheet_name, range_name)
        
        scale = USABLE_WIDTH / w
        target_w = USABLE_WIDTH
        target_h = h * scale
        
        total_h += target_h + 160  
        images_data.append((img_reader, target_w, target_h, description))
        
    HEADER_HEIGHT = 150
    PAGE_HEIGHT = total_h + (MARGIN * 2) + HEADER_HEIGHT + 50 # Add extra bottom margin just in case
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))
    
    # Draw Header Background
    c.setFillColorRGB(0.08, 0.15, 0.36) # Professional Dark Blue
    c.rect(0, PAGE_HEIGHT - HEADER_HEIGHT, PAGE_WIDTH, HEADER_HEIGHT, fill=True, stroke=False)
    
    # Draw Header Text
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 48)
    c.drawString(MARGIN, PAGE_HEIGHT - HEADER_HEIGHT / 2 + 10, "PW Online - Vishwas Diwas Analytics Report")
    
    c.setFont("Helvetica", 24)
    c.drawString(MARGIN, PAGE_HEIGHT - HEADER_HEIGHT / 2 - 30, f"Generated on: {TODAY}")
    
    current_y = PAGE_HEIGHT - HEADER_HEIGHT - MARGIN

    for img_reader, target_w, target_h, description in images_data:
        clean_desc = description.lstrip('#').strip()
        
        # Section Header Box
        c.setFillColorRGB(0.96, 0.96, 0.98) # Very light blue-grey
        c.roundRect(MARGIN, current_y - 60, USABLE_WIDTH, 60, 10, fill=True, stroke=False)
        
        # Section Header Text
        c.setFillColorRGB(0.1, 0.1, 0.1)
        c.setFont("Helvetica-Bold", 30)
        # Vertically center text in the box
        c.drawString(MARGIN + 20, current_y - 42, clean_desc)
        
        current_y -= 80 # Move past header box + 20px padding
        current_y -= target_h # Move to bottom of image placement
        
        # Draw image border
        c.setStrokeColorRGB(0.85, 0.85, 0.85)
        c.setLineWidth(2)
        c.rect(MARGIN - 2, current_y - 2, target_w + 4, target_h + 4, fill=False, stroke=True)
        
        # Draw Image
        c.drawImage(img_reader, MARGIN, current_y, width=target_w, height=target_h, preserveAspectRatio=True, mask='auto')
        
        current_y -= 80 # Padding between bottom of image and start of next section

    # Footer
    c.setFillColorRGB(0.6, 0.6, 0.6)
    c.setFont("Helvetica", 20)
    c.drawCentredString(PAGE_WIDTH / 2.0, MARGIN / 2, "CONFIDENTIAL - Internal Use Only")

    c.save()
    buffer.seek(0)
    print("✅ FINAL: Large, UHD, Single Page Dynamic PDF Generated")
    return buffer

def upload_to_r2(pdf_buffer):
    s3 = boto3.client(
        service_name="s3",
        endpoint_url=f"https://{R2_ACCOUNT_ID}.r2.cloudflarestorage.com",
        aws_access_key_id=R2_ACCESS_KEY,
        aws_secret_access_key=R2_SECRET_KEY,
        region_name="auto"
    )

    s3.put_object(
        Bucket=R2_BUCKET_NAME,
        Key=FILE_NAME,
        Body=pdf_buffer.read(),
        ContentType="application/pdf"
    )

    public_url = f"{R2_PUBLIC_BASE}/{FILE_NAME}"

    print("✅ Uploaded to Cloudflare R2")
    print("🔗 Public URL:", public_url)

    return public_url

def send_to_aisensy(url):
    endpoint = "https://backend.aisensy.com/campaign/t1/api/v2"

    for dest in DESTINATIONS:
        payload = {
            "apiKey": AISENSY_API_KEY,
            "campaignName": CAMPAIGN_NAME,
            "destination": dest,
            "userName": "PW Online- Analytics",
            "templateParams": [TODAY],
            "source": "r2-centered",
            "media": {
                "url": url,
                "filename": FILE_NAME
            }
        }

        response = requests.post(
            endpoint,
            json=payload,
            headers={"Content-Type": "application/json"}
        )

        print(f"📱 Sent to WhatsApp ({dest}):", response.status_code, response.text)

if __name__ == "__main__":
    try:
        missing_vars = []
        for v in [
            "GOOGLE_CREDENTIALS_JSON", "R2_ACCOUNT_ID", "R2_ACCESS_KEY", "R2_SECRET_KEY", "R2_BUCKET_NAME", "R2_PUBLIC_BASE",
            "AISENSY_API_KEY", "SHEET_ID", "DESTINATIONS", "CAMPAIGN_NAME"
        ]:
            if not os.environ.get(v):
                missing_vars.append(v)
        
        if missing_vars:
            raise Exception(f"Missing required environment variables: {', '.join(missing_vars)}")
            
        print("🚀 Starting Dynamically Scaled PDF Generation...")
        pdf_buffer = generate_dynamic_single_page_clean()

        public_pdf_url = upload_to_r2(pdf_buffer)

        send_to_aisensy(public_pdf_url)

        print("🎉 Successfully completed dynamic PDF automation via GitHub Actions!")
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"❌ Error occurred: {e}")
        exit(1)

