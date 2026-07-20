import os
import shutil
import subprocess
from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
from pdf2image import convert_from_path

app = FastAPI()

@app.post("/convert-ppt-to-png/")
async def convert_ppt_to_png(file: UploadFile = File(...)):
    temp_dir = "temp_files"
    os.makedirs(temp_dir, exist_ok=True)
    
    # 1. 업로드된 파워포인트 파일 저장
    file_path = os.path.join(temp_dir, file.filename)
    with open(file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    pdf_path = os.path.join(temp_dir, "output.pdf")
    
    # 2. LibreOffice를 이용해 PPTX -> PDF 변환
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf", 
        file_path, "--outdir", temp_dir
    ], check=True)
    
    # 3. PDF의 각 페이지를 PNG 이미지로 변환
    images = convert_from_path(pdf_path)
    image_paths = []
    
    for i, image in enumerate(images):
        img_path = os.path.join(temp_dir, f"slide_{i + 1}.png")
        image.save(img_path, "PNG")
        image_paths.append(img_path)
        
    # 4. 첫 번째 슬라이드 이미지를 응답으로 반환
    return FileResponse(image_paths[0], media_type="image/png", filename="slide_1.png")
