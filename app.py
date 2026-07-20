import os
import subprocess
import tempfile
import streamlit as st
from pdf2image import convert_from_path

st.title("파워포인트(PPTX)를 PNG 이미지로 변환")
st.write("파워포인트 파일을 업로드하면 각 슬라이드를 PNG 이미지로 변환해 드립니다.")

# 파일 업로더
uploaded_file = st.file_uploader(
    "PPTX 파일을 선택하세요", type=["pptx", "ppt"]
)

if uploaded_file is not None:
  if st.button("변환 시작"):
    with st.spinner("변환 중입니다. 잠시만 기다려주세요..."):
      try:
        # 1. 임시 디렉토리 생성
        os.makedirs("temp_files", exist_ok=True)

        # 2. 업로드된 파일을 임시 경로에 저장
        input_path = os.path.join("temp_files", uploaded_file.name)
        with open(input_path, "wb") as f:
          f.write(uploaded_file.getbuffer())

        # 3. LibreOffice를 사용해 PPTX -> PDF 변환
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                input_path,
                "--outdir",
                "temp_files",
            ],
            check=True,
        )

        # 4. LibreOffice가 생성한 PDF 파일명을 'output.pdf'로 통일 (이름 매칭 처리)
        original_pdf_name = os.path.splitext(uploaded_file.name)[0] + ".pdf"
        generated_pdf_path = os.path.join("temp_files", original_pdf_name)
        target_pdf_path = os.path.join("temp_files", "output.pdf")

        if os.path.exists(generated_pdf_path):
          if os.path.exists(target_pdf_path):
            os.remove(target_pdf_path)
          os.rename(generated_pdf_path, target_pdf_path)
        else:
          raise FileNotFoundError(
              f"변환된 PDF 파일을 찾을 수 없습니다: {generated_pdf_path}"
          )

        # 5. pdf2image를 사용해 PDF -> PNG 이미지로 변환
        images = convert_from_path(target_pdf_path, dpi=300)

        st.success(f"총 {len(images)}개의 슬라이드가 변환되었습니다!")

        # 6. 변환된 이미지 화면 출력 및 다운로드 제공
        for i, image in enumerate(images):
          st.image(
              image,
              caption=f"슬라이드 {i + 1}",
              use_container_width=True,
          )

          # 이미지 다운로드 버튼
          img_path = os.path.join("temp_files", f"slide_{i + 1}.png")
          image.save(img_path, "PNG")

          with open(img_path, "rb") as img_file:
            st.download_button(
                label=f"슬라이드 {i + 1} 다운로드 (PNG)",
                data=img_file,
                file_name=f"slide_{i + 1}.png",
                mime="image/png",
                key=f"download_{i}",
            )

      except Exception as e:
        st.error(f"변환 중 오류가 발생했습니다: {e}")
