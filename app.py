import streamlit as st
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
import io
import zipfile

st.title("학생 이미지 추출 및 파일명 자동 지정 변환기")
st.write("PPTX 파일을 업로드하면 각 슬라이드의 상단 텍스트를 파일명으로 하여 이미지들을 하나의 투명 PNG로 합쳐서 추출합니다.")

uploaded_file = st.file_uploader("파워포인트 파일(.pptx)을 업로드하세요", type=["pptx"])

# 재귀적으로 도형 내부를 돌며 이미지와 위치 정보를 함께 수집하는 함수
def extract_images_with_position(shapes, image_list):
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            extract_images_with_position(shape.shapes, image_list)
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            # 이미지와 슬라이드 내 좌표(left, top)를 함께 저장
            image_list.append({
                'image': shape.image,
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height
            })

if uploaded_file is not None:
    with st.spinner("이미지를 병합하는 중입니다..."):
        prs = Presentation(uploaded_file)
        
        # 슬라이드 가로/세로 크기 (기본 16:9 비율 기준 예시, pptx 자체 크기 활용 가능)
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for slide_index, slide in enumerate(prs.slides):
                # 1. 상단 텍스트를 파일명으로 지정
                slide_title = f"slide_{slide_index + 1}"
                for shape in slide.shapes:
                    if shape.has_text_frame and shape.text.strip():
                        clean_text = shape.text.strip().split('\n')[0]
                        for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
                            clean_text = clean_text.replace(ch, '')
                        if clean_text:
                            slide_title = clean_text
                            break
                
                # 2. 슬라이드 내 이미지 및 위치 수집
                slide_images = []
                extract_images_with_position(slide.shapes, slide_images)
                
                if not slide_images:
                    continue
                
                try:
                    # 3. 투명 배경을 가진 빈 캔버스 생성 (슬라이드 크기 비율에 맞춰 고해상도 설정)
                    canvas_width = 1920
                    canvas_height = int(1920 * (slide_height / slide_width))
                    base_canvas = Image.new("RGBA", (canvas_width, canvas_height), (0, 0, 0, 0))
                    
                    # 4. 수집된 각 이미지를 슬라이드 좌표 비율에 맞추어 캔버스에 붙여넣기
                    for item in slide_images:
                        img_stream = io.BytesIO(item['image'].blob)
                        pil_img = Image.open(img_stream).convert("RGBA")
                        
                        # PPT 좌표를 캔버스 픽셀 좌표로 환산
                        x = int(item['left'] / slide_width * canvas_width)
                        y = int(item['top'] / slide_height * canvas_height)
                        w = int(item['width'] / slide_width * canvas_width)
                        h = int(item['height'] / slide_height * canvas_height)
                        
                        # 크기 조절
                        resized_img = pil_img.resize((w, h), Image.Resampling.LANCZOS)
                        
                        # 투명도를 유지하며 빈 캔버스에 합성
                        base_canvas.paste(resized_img, (x, y), resized_img)
                    
                    # 5. 완성된 병합 이미지를 바이트로 변환하여 ZIP에 저장
                    img_byte_arr = io.BytesIO()
                    base_canvas.save(img_byte_arr, format="PNG")
                    img_byte_arr.seek(0)
                    
                    file_name = f"{slide_title}.png"
                    zip_file.writestr(file_name, img_byte_arr.read())
                    
                except Exception as e:
                    print(f"슬라이드 {slide_index + 1} 병합 실패: {e}")
                        
        zip_buffer.seek(0)
        
        st.success("모든 슬라이드 이미지 병합이 완료되었습니다!")
        st.download_button(
            label="병합된 이미지 다운로드 (ZIP)",
            data=zip_buffer,
            file_name="merged_images.zip",
            mime="application/zip"
        )
