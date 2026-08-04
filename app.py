import streamlit as st
import os
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
import io
import zipfile

st.title("학생 이미지 추출 및 파일명 자동 지정 변환기")
st.write("PPTX 파일을 업로드하면 각 슬라이드의 상단 텍스트를 파일명으로 하여 이미지들을 개별 PNG로 추출합니다.")

uploaded_file = st.file_uploader("파워포인트 파일(.pptx)을 업로드하세요", type=["pptx"])

# 재귀적으로 도형(그룹 포함) 내부를 돌며 이미지를 찾는 함수
def extract_images_from_shapes(shapes, image_list):
    for shape in shapes:
        # 그룹화된 도형인 경우 내부 도형들을 다시 탐색
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            extract_images_from_shapes(shape.shapes, image_list)
        # 그림(Picture) 객체인 경우
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            image_list.append(shape.image)

if uploaded_file is not None:
    with st.spinner("이미지를 추출하는 중입니다..."):
        prs = Presentation(uploaded_file)
        
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
                
                # 2. 슬라이드 내 이미지 수집 (그룹 내부 포함)
                slide_images = []
                extract_images_from_shapes(slide.shapes, slide_images)
                
                # 3. 수집된 이미지들을 파일로 변환하여 ZIP에 추가
                for img_counter, image in enumerate(slide_images, start=1):
                    image_bytes = image.blob
                    
                    if img_counter == 1:
                        file_name = f"{slide_title}.png"
                    else:
                        file_name = f"{slide_title}_{img_counter}.png"
                    
                    try:
                        image_stream = io.BytesIO(image_bytes)
                        img = Image.open(image_stream)
                        
                        img_byte_arr = io.BytesIO()
                        img.save(img_byte_arr, format="PNG")
                        img_byte_arr.seek(0)
                        
                        zip_file.writestr(file_name, img_byte_arr.read())
                    except Exception as e:
                        print(f"이미지 처리 실패: {e}")
                        
        zip_buffer.seek(0)
        
        st.success("모든 이미지 추출이 완료되었습니다!")
        st.download_button(
            label="추출된 이미지 다운로드 (ZIP)",
            data=zip_buffer,
            file_name="extracted_images.zip",
            mime="application/zip"
        )
