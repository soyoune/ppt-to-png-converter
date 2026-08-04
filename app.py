import io
import zipfile
import streamlit as st
from pptx import Presentation
from PIL import Image

st.title("학생 이미지 추출 및 파일명 자동 지정 변환기")
st.write("각 슬라이드의 상단 텍스트를 파일명으로 지정하고 이미지를 추출합니다.")

uploaded_file = st.file_uploader("파워포인트(PPTX) 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    if st.button("변환 시작"):
        with st.spinner("이미지와 파일명을 추출하는 중입니다..."):
            
            prs = Presentation(uploaded_file)
            zip_buffer = io.BytesIO()
            success_count = 0
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                
                for i, slide in enumerate(prs.slides):
                    student_name = f"학생_{i+1}"
                    slide_image_blob = None
                    
                    # 1. 텍스트와 이미지 요소 추출
                    for shape in slide.shapes:
                        # 텍스트 상자에서 학생 이름(파일명) 탐색
                        if shape.has_text_frame:
                            text = shape.text_frame.text.strip()
                            if text:
                                safe_text = "".join(c for c in text if c.isalnum() or c in (' ', '_', '-', '(', ')')).strip()
                                if safe_text:
                                    if student_name.startswith("학생_"):
                                        student_name = safe_text
                        
                        # 슬라이드 내 이미지 객체 추출
                        if shape.shape_type == 13 or hasattr(shape, "image"):
                            try:
                                slide_image_blob = shape.image.blob
                            except Exception:
                                pass
                    
                    # 2. 이미지 처리 및 저장
                    if slide_image_blob:
                        try:
                            image = Image.open(io.BytesIO(slide_image_blob))
                            
                            # RGBA 모드로 변환 (투명도 지원 포맷 유지)
                            image = image.convert("RGBA")
                            
                            img_byte_arr = io.BytesIO()
                            image.save(img_byte_arr, format='PNG')
                            img_byte_arr = img_byte_arr.getvalue()
                            
                            file_name = f"{student_name}.png"
                            zip_file.writestr(file_name, img_byte_arr)
                            success_count += 1
                            
                        except Exception as e:
                            st.warning(f"슬라이드 {i+1} 처리 중 오류: {e}")

            zip_buffer.seek(0)
            
            if success_count > 0:
                st.success(f"총 {success_count}개의 이미지 변환 완료!")
                st.download_button(
                    label="이미지 ZIP으로 다운로드",
                    data=zip_buffer,
                    file_name="students_images.zip",
                    mime="application/zip"
                )
            else:
                st.error("슬라이드에서 이미지를 찾지 못했습니다.")
