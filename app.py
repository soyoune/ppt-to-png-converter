import io
import zipfile
import streamlit as st
from pptx import Presentation
from PIL import Image
from rembg import remove

st.title("학생 이미지 추출 및 투명화 변환기")
st.write("각 슬라이드의 상단 텍스트를 파일명으로 가져오고, 이미지 배경을 투명하게 제거하여 일괄 다운로드합니다.")

uploaded_file = st.file_uploader("파워포인트(PPTX) 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    if st.button("변환 시작"):
        with st.spinner("텍스트를 추출하고 이미지 배경을 투명하게 만드는 중입니다..."):
            
            prs = Presentation(uploaded_file)
            zip_buffer = io.BytesIO()
            success_count = 0
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                
                for i, slide in enumerate(prs.slides):
                    student_name = f"학생_{i+1}"
                    slide_image_blob = None
                    
                    # 1. 슬라이드 내의 모든 도형(텍스트 상자 및 이미지)을 순회하며 데이터 수집
                    for shape in slide.shapes:
                        # 텍스트 상자에서 이름 추출 (상단에 위치한 텍스트 우선 탐색)
                        if shape.has_text_frame:
                            text = shape.text_frame.text.strip()
                            if text:
                                # 파일명으로 사용할 수 없는 특수문자 제거 및 정제
                                safe_text = "".join(c for c in text if c.isalnum() or c in (' ', '_', '-', '(', ')')).strip()
                                if safe_text:
                                    # 첫 번째로 발견된 유효한 텍스트를 학생 이름으로 지정
                                    if student_name.startswith("학생_"):
                                        student_name = safe_text
                        
                        # 슬라이드에 포함된 이미지(그림) 객체 추출
                        if shape.shape_type == 13 or hasattr(shape, "image"):
                            try:
                                slide_image_blob = shape.image.blob
                            except Exception:
                                pass
                    
                    # 2. 이미지가 존재할 경우에만 배경 투명화 및 파일 저장 진행
                    if slide_image_blob:
                        try:
                            # 바이트 데이터를 PIL 이미지로 변환
                            image = Image.open(io.BytesIO(slide_image_blob))
                            
                            # 투명도를 지원하는 RGBA 모드로 변환
                            image = image.convert("RGBA")
                            
                            # rembg 라이브러리를 통해 배경을 완전히 제거하고 투명하게 처리
                            output_image = remove(image)
                            
                            # PNG 형식의 바이트로 변환
                            img_byte_arr = io.BytesIO()
                            output_image.save(img_byte_arr, format='PNG')
                            img_byte_arr = img_byte_arr.getvalue()
                            
                            # 중복 파일명 방지 처리 (이름이 같을 경우 번호 붙이기)
                            file_name = f"{student_name}.png"
                            
                            # ZIP 파일에 압축 저장
                            zip_file.writestr(file_name, img_byte_arr)
                            success_count += 1
                            
                        except Exception as e:
                            st.warning(f"슬라이드 {i+1} 처리 중 오류 발생: {e}")

            zip_buffer.seek(0)
            
            if success_count > 0:
                st.success(f"총 {success_count}개의 학생 이미지 변환 및 배경 투명화 완료!")
                
                st.download_button(
                    label="투명 PNG 이미지들 ZIP으로 다운로드",
                    data=zip_buffer,
                    file_name="transparent_students_images.zip",
                    mime="application/zip"
                )
            else:
                st.error("슬라이드에서 이미지 개체를 찾지 못했습니다. 파워포인트 내 이미지가 '그림' 형태로 제대로 삽입되어 있는지 확인해 주세요.")
