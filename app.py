import streamlit as st
import os
from pptx import Presentation
from PIL import Image
import io
import zipfile

st.title("학생 이미지 추출 및 파일명 자동 지정 변환기")
st.write("PPTX 파일을 업로드하면 각 슬라이드의 상단 텍스트를 파일명으로 하여 이미지들을 개별 PNG로 추출합니다.")

uploaded_file = st.file_uploader("파워포인트 파일(.pptx)을 업로드하세요", type=["pptx"])

if uploaded_file is not None:
    with st.spinner("이미지를 추출하는 중입니다..."):
        prs = Presentation(uploaded_file)
        
        # 임시 결과 저장용 바이트 버퍼 생성
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
                
                # 2. 슬라이드 내 이미지 추출
                img_counter = 1
                for shape in slide.shapes:
                    if shape.shape_type == 1:  # 그림 객체
                        image = shape.image
                        image_bytes = image.blob
                        
                        if img_counter == 1:
                            file_name = f"{slide_title}.png"
                        else:
                            file_name = f"{slide_title}_{img_counter}.png"
                        
                        try:
                            image_stream = io.BytesIO(image_bytes)
                            img = Image.open(image_stream)
                            
                            # PNG 포맷으로 변환용 바이트 버퍼
                            img_byte_arr = io.BytesIO()
                            img.save(img_byte_arr, format="PNG")
                            img_byte_arr.seek(0)
                            
                            # ZIP 파일에 추가
                            zip_file.writestr(file_name, img_byte_arr.read())
                        except Exception as e:
                            print(f"이미지 처리 실패: {e}")
                            
                        img_counter += 1
                        
        zip_buffer.seek(0)
        
        st.success("모든 이미지 추출이 완료되었습니다!")
        st.download_button(
            label="추출된 이미지 다운로드 (ZIP)",
            data=zip_buffer,
            file_name="extracted_images.zip",
            mime="application/zip"
        )
