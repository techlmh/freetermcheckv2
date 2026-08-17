# 📚 중학교 자유학기 운영계획서 자동 검토 시스템 (Streamlit WebApp)

본 프로젝트는 교육청 표준 서식으로 작성된 **중학교 자유학기 운영계획서(Excel)**를 자동으로 분석하고, 교육 지침 준수 여부를 검증하는 Streamlit 기반 웹 애플리케이션입니다.

## 🚀 주요 기능

- **일괄 업로드 및 검토**: 여러 학교의 엑셀 파일을 동시에 업로드하여 개별 탭으로 결과 확인

- **자동 검증 로직**: 운영 시수, 활동 영역별 정합성, 예산 산출 근거 등 10여 가지 항목 자동 점검

- **시각적 피드백**: 보완이 필요한 항목(`FAIL`)에 대해 빨간색 배경 강조 및 상세 불일치 사유 표시

- **결과 리포트 다운로드**: 모든 학교의 검토 결과를 시트별로 정리하고 서식이 적용된 통합 엑셀 파일 제공

- **세션 관리**: 업로드 파일 일괄 초기화 기능을 통한 편리한 반복 작업 지원

## 🛠 기술 스택

- **Language**: Python 3.11+

- **Framework**: Streamlit (Web UI & State Management)

- **Data Processing**: Pandas (Result Table), Openpyxl (Excel Parsing & Styling)

- **Deployment**: GitHub & Streamlit Community Cloud

## 📋 설치 및 실행 방법 (Local)

1. 저장소 클론:

```bash
git clone https://github.com/[사용자ID]/[저장소명].git
cd [저장소명]
```

1. 필수 라이브러리 설치:

```bash
pip install -r requirements.txt
```

1. 웹앱 실행:

```bash
streamlit run app.py
```

## 📂 파일 구조

- `app.py`: 웹앱 메인 소스 코드 (UI, 파싱, 검증 로직 통합 )

- `requirements.txt`: 배포에 필요한 파이썬 의존성 목록

- `README.md`: 프로젝트 개요 및 설치 가이드

- `Handover_Guide.md`: 상세 기술 설계 및 업데이트 가이드

## ⚖️ 라이선스

본 프로그램은 학교 행정 업무 경감을 위해 제작되었습니다.데이터 파싱 로직은 서울특별시교육청 2026학년도 자유학기 운영계획서 표준 서식을 기준으로 합니다.
