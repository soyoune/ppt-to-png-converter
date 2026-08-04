import os
from pptx import Presentation
from PIL import Image
import io

def extract_images_from_pptx(pptx_path, output_dir="extracted_images"):
    # 출력 폴더 생성
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    prs = Presentation(pptx_path)
    
    for slide_index, slide in enumerate(prs.slides):
        # 1. 상단 텍스트(또는 첫 번째 텍스트 상자)를 찾아서 파일명으로 지정
        slide_title = f"slide_{slide_index + 1}"
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip():
                # 줄바꿈이나 공백을 제거하여 파일명으로 안전한 문자열 생성
                clean_text = shape.text.strip().split('\n')[0]
                # 파일명으로 사용할 수 없는 특수문자 제거
                for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
                    clean_text = clean_text.replace(ch, '')
                if clean_text:
                    slide_title = clean_text
                    break
        
        # 2. 슬라이드 내의 모든 도형을 순회하며 이미지 추출
        img_counter = 1
        for shape in slide.shapes:
            # 그림(Picture) 객체인 경우
            if shape.shape_type == 1: # 1은 MSO_SHAPE_TYPE.PICTURE 에 해당
                image = shape.image
                image_bytes = image.blob
                image_ext = image.ext  # 원본 확장자 (png, jpeg 등)
                
                # 파일명 설정 (슬라이드 내에 이미지가 여러 개일 경우 번호 부여)
                if img_counter == 1:
                    file_name = f"{slide_title}.png"
                else:
                    file_name = f"{slide_title}_{img_counter}.png"
                
                file_path = os.path.join(output_dir, file_name)
                
                # 이미지 데이터를 PIL을 통해 처리하여 PNG로 저장 (투명도 유지)
                try:
                    image_stream = io.BytesIO(image_bytes)
                    img = Image.open(image_stream)
                    
                    # 투명 배경(RGBA 등)이 있는 경우 그대로 PNG로 저장
                    img.save(file_path, "PNG")
                    print(f"추출 완료: {file_path}")
                except Exception as e:
                    print(f"이미지 저장 실패 ({file_name}): {e}")
                
                img_counter += 1

# 사용 예시
# pptx_path = "당신의_파이썬파일경로.pptx"
# extract_images_from_pptx(pptx_path)
