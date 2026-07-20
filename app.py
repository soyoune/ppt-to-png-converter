import os
import shutil
import subprocess
import streamlit as st
from pdf2image import convert_from_path

st.title("파워포인트(PPTX)를 PNG 이미지로 변환")
st.write("파워포인트 파일을 업로드하면 각 슬라이드를 PNG 이미지로 변환해 드립니다.")

# 파일 업로드 위젯
uploaded_file = st.file_uploader("PPTX 파일을 선택하세요", type=["pptx"])

if uploaded_file is not None:
    if st.button("변환 시작"):
        with st.spinner("변환 중입니다... 잠시만 기다려주세요!"):
            temp_dir = "temp_files"
            os.makedirs(temp_dir, exist_ok=True)
            
            # 업로드된 파일 저장
            file_path = os.path.join(temp_dir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
                
            pdf_path = os.path.join(temp_dir, "output.pdf")
            
            try:
                # LibreOffice를 이용해 PPTX -> PDF 변환
                subprocess.run([
                    "libreoffice", "--headless", "--convert-to", "pdf", 
                    file_path, "--outdir", temp_dir
                ], check=True)
                
                # PDF -> PNG 변환
                images = convert_from_path(pdf_path)
                
                st.success(f"총 {len(images)}개의 슬라이드가 변환되었습니다!")
                
                # 슬라이드 이미지 화면에 표시 및 다운로드 제공
                for i, image in enumerate(images):
                    img_path = os.path.join(temp_dir, f"slide_{i + 1}.png")
                    image.save(img_path, "PNG")
                    
                    st.image(image, caption=f"슬라이드 {i + 1}", use_container_width=True)
                    
                    with open(img_path, "rb") as img_file:
                        st.download_button(
                            label=f"슬라이드 {i + 1} 다운로드",
                            data=img_file,
                            file_name=f"slide_{i + 1}.png",
                            mime="image/png"
                        )
            except Exception as e:
                st.error(f"변환 중 오류가 발생했습니다: {e}")
