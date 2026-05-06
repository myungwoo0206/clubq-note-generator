import streamlit as st
import subprocess
import sys
import os
from PIL import Image

# -----------------------------
# Playwright Chromium 설치
# -----------------------------
def ensure_playwright():
    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=False
    )

ensure_playwright()

# -----------------------------
# Streamlit 페이지 설정
# -----------------------------
st.set_page_config(
    page_title="ClubQ NOTE Generator",
    layout="centered"
)

# -----------------------------
# 제목
# -----------------------------
st.title("📊 ClubQ NOTE 이미지 생성기")

st.write(
    "엑셀 파일을 업로드하면 NOTE 이미지를 자동 생성합니다."
)

# -----------------------------
# 엑셀 업로드
# -----------------------------
uploaded_file = st.file_uploader(
    "data.xlsx 업로드",
    type=["xlsx"]
)

# -----------------------------
# 업로드 성공 시
# -----------------------------
if uploaded_file:

    # 업로드 파일 저장
    with open("data.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success("✅ 엑셀 업로드 완료")

    # -----------------------------
    # 이미지 생성 버튼
    # -----------------------------
    if st.button("이미지 생성하기"):

        with st.spinner("이미지 생성 중입니다..."):

            result = subprocess.run(
                [sys.executable, "generate.py"],
                capture_output=True,
                text=True
            )

        # -----------------------------
        # 성공
        # -----------------------------
        if result.returncode == 0:

            st.success("✅ 이미지 생성 완료")

            image_path = "clubq_note_final.png"

            if os.path.exists(image_path):

                image = Image.open(image_path)

                st.image(
                    image,
                    caption="생성된 ClubQ NOTE",
                    use_container_width=True
                )

                with open(image_path, "rb") as file:

                    st.download_button(
                        label="📥 이미지 다운로드",
                        data=file,
                        file_name="clubq_note_final.png",
                        mime="image/png"
                    )

            else:
                st.error("❌ 이미지 파일을 찾을 수 없습니다.")

        # -----------------------------
        # 실패
        # -----------------------------
        else:

            st.error("❌ 이미지 생성 실패")

            st.code(result.stderr)
